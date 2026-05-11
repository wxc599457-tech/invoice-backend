# -*- coding: utf-8 -*-
"""
开票汇总助手 —— 后端 Flask 应用
功能：
1. 接收前端上传的 Excel 文件 (.xls / .xlsx)
2. 自动识别 “开票名称 / 重量 / 金额” 三列
3. 清理开票名称（去掉中英文括号及其中内容）
4. 自动识别 kg / g 并统一换算为 kg
5. 按清理后的开票名称分组，汇总总重量、总金额
6. 生成 “汇总结果.xlsx” 供前端下载，并保存到 backend/output/ 目录
"""

import io
import os
import re
import uuid
import logging
from datetime import datetime

import pandas as pd
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# ---------------------------------------------------------------------------
# 基础配置


# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"xls", "xlsx"}
MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 单次最大上传 32MB

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
CORS(app)  # 允许前端跨域访问
from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os

# -------------------------
# 健康检查接口
# -------------------------
@app.route("/api/health", methods=["GET"])
def health_check():
    return jsonify({"status": "ok"})

# 你原来的上传、汇总 Excel 的代码保持不变
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s: %(message)s")
log = logging.getLogger("invoice-app")


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def allowed_file(filename: str) -> bool:
    """判断文件后缀是否合法"""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def clean_invoice_name(name) -> str:
    """
    清理开票名称：
    去除中文括号（）和英文括号()以及括号内的全部内容，
    只保留括号前的部分。
    示例：
        "杭州化工（1月）"  -> "杭州化工"
        "杭州化工(补单)"   -> "杭州化工"
        "ABC Corp (2024)" -> "ABC Corp"
    """
    if name is None:
        return ""
    text = str(name).strip()
    # 同时匹配中文与英文括号及其内部任意字符
    cleaned = re.split(r"[（(]", text, maxsplit=1)[0]
    return cleaned.strip()


# 重量字符串解析正则：捕获数字（可含小数）和单位（kg / g / 千克 / 克）
_WEIGHT_PATTERN = re.compile(
    r"(-?\d+(?:\.\d+)?)\s*(kg|g|千克|克|公斤)?",
    re.IGNORECASE,
)


def convert_weight_to_kg(value) -> float:
    """
    将重量字段统一换算为 kg。
    支持的输入：
        - 数值类型（默认按 kg 处理）
        - 字符串（如 "500g"、"1kg"、"2.5 千克"、"800 克"）
    示例：
        500g  -> 0.5
        1kg   -> 1.0
        2500g -> 2.5
    """
    if value is None:
        return 0.0

    # 数值类型：默认按 kg
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return 0.0
        return float(value)

    text = str(value).strip().lower().replace(",", "")
    if text == "" or text == "nan":
        return 0.0

    match = _WEIGHT_PATTERN.search(text)
    if not match:
        # 实在无法解析则返回 0，避免整个流程崩溃
        log.warning("无法解析的重量字段：%s", value)
        return 0.0

    number = float(match.group(1))
    unit = (match.group(2) or "kg").lower()

    if unit in ("g", "克"):
        return number / 1000.0
    # kg / 千克 / 公斤 / 缺省单位
    return number


def find_column(df_columns, candidates):
    """
    在 DataFrame 的列名中模糊查找目标列。
    candidates 为可能的关键字列表，命中其中任何一个即认为匹配。
    """
    for col in df_columns:
        col_str = str(col).strip()
        for key in candidates:
            if key in col_str:
                return col
    return None


# ---------------------------------------------------------------------------
# 核心处理逻辑
# ---------------------------------------------------------------------------
def process_excel(file_stream, original_filename: str):
    """
    读取上传的 Excel 文件并完成全部数据处理。
    返回 (summary_df, output_path)。
    任何业务异常会以 ValueError 形式抛出，并由路由层统一返回给前端。
    """
    # 根据文件后缀选择引擎
    ext = original_filename.rsplit(".", 1)[1].lower()
    try:
        if ext == "xlsx":
            df = pd.read_excel(file_stream, engine="openpyxl")
        else:  # xls
            # xlrd 仅在安装且为 1.2.0 时支持 xls；这里依然尝试默认引擎
            df = pd.read_excel(file_stream)
    except Exception as exc:
        raise ValueError(f"Excel 读取失败，请确认文件格式正确：{exc}") from exc

    if df is None or df.empty:
        raise ValueError("Excel 内容为空，请检查后重新上传。")

    # 自动识别三列
    col_name = find_column(df.columns, ["开票名称", "名称", "客户", "公司"])
    col_weight = find_column(df.columns, ["重量", "weight", "净重"])
    col_amount = find_column(df.columns, ["金额", "amount", "价格", "总价"])

    missing = []
    if col_name is None:
        missing.append("开票名称")
    if col_weight is None:
        missing.append("重量")
    if col_amount is None:
        missing.append("金额")
    if missing:
        raise ValueError("Excel 缺少必要列：" + "、".join(missing))

    work = pd.DataFrame({
        "开票名称": df[col_name].apply(clean_invoice_name),
        "重量(kg)": df[col_weight].apply(convert_weight_to_kg),
        "金额": pd.to_numeric(df[col_amount], errors="coerce").fillna(0.0),
    })

    # 去掉名称为空的行
    work = work[work["开票名称"].astype(str).str.len() > 0]

    if work.empty:
        raise ValueError("有效数据为空，请检查 Excel 内容。")

    summary = (
        work.groupby("开票名称", as_index=False)
        .agg({"重量(kg)": "sum", "金额": "sum"})
        .rename(columns={"重量(kg)": "总重量(kg)", "金额": "总金额"})
        .sort_values(by="总金额", ascending=False)
        .reset_index(drop=True)
    )
    # 数值四舍五入，避免浮点误差
    summary["总重量(kg)"] = summary["总重量(kg)"].round(3)
    summary["总金额"] = summary["总金额"].round(2)

    # 写出文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_id = f"{timestamp}_{uuid.uuid4().hex[:6]}"
    output_filename = f"汇总结果_{file_id}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="汇总结果")
        worksheet = writer.sheets["汇总结果"]
        # 简单的列宽自适应
        for idx, column in enumerate(summary.columns, start=1):
            max_len = max(
                [len(str(column))] + [len(str(v)) for v in summary[column].tolist()]
            )
            worksheet.column_dimensions[chr(64 + idx)].width = min(max_len + 4, 40)

    return summary, output_filename


# ---------------------------------------------------------------------------
# 路由
# ---------------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    """健康检查接口"""
    return jsonify({"status": "ok", "service": "开票汇总助手"})


@app.route("/api/upload", methods=["POST"])
def upload():
    """
    接收前端上传的 Excel 文件并返回汇总结果。
    返回：
        {
            "success": True,
            "summary": [...],     # 汇总表数据
            "file": "汇总结果_xxx.xlsx", # 下载用文件名
            "rows": 12            # 汇总行数
        }
    """
    if "file" not in request.files:
        return jsonify({"success": False, "message": "未检测到上传的文件，请重新选择。"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "message": "文件名为空，请重新选择文件。"}), 400

    if not allowed_file(file.filename):
        return jsonify({"success": False, "message": "仅支持 .xls / .xlsx 格式的 Excel 文件。"}), 400

    safe_name = secure_filename(file.filename) or "upload.xlsx"
    saved_upload_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex[:8]}_{safe_name}")
    try:
        file.save(saved_upload_path)
    except Exception as exc:
        log.exception("保存上传文件失败")
        return jsonify({"success": False, "message": f"文件保存失败：{exc}"}), 500

    try:
        with open(saved_upload_path, "rb") as f:
            buf = io.BytesIO(f.read())
        summary_df, output_filename = process_excel(buf, file.filename)
    except ValueError as ve:
        log.warning("业务校验失败：%s", ve)
        return jsonify({"success": False, "message": str(ve)}), 400
    except Exception as exc:
        log.exception("处理 Excel 时发生未知错误")
        return jsonify({"success": False, "message": f"处理失败：{exc}"}), 500

    summary_records = summary_df.to_dict(orient="records")
    return jsonify({
        "success": True,
        "message": f"汇总成功，共 {len(summary_records)} 条记录。",
        "summary": summary_records,
        "file": output_filename,
        "rows": len(summary_records),
    })


@app.route("/api/download/<path:filename>", methods=["GET"])
def download(filename: str):
    """根据文件名下载已生成的汇总结果"""
    safe_name = os.path.basename(filename)
    target = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(target):
        return jsonify({"success": False, "message": "文件不存在或已被清理。"}), 404

    return send_file(
        target,
        as_attachment=True,
        download_name="汇总结果.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(_):
    return jsonify({"success": False, "message": "文件过大，单次上传请控制在 32MB 以内。"}), 413


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print(" 开票汇总助手 - 后端服务")
    print(" 默认地址: http://127.0.0.1:5000")
    print(" 健康检查: http://127.0.0.1:5000/api/health")
    print("=" * 60)
    app.run(host="127.0.0.1", port=5000, debug=False)
