# app.py
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
import logging

# ----------------------------
# Flask app 初始化
# ----------------------------
app = Flask(__name__)
CORS(app)  # 允许跨域请求

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("invoice-app")

# ----------------------------
# 健康检查接口
# ----------------------------
@app.route("/api/health", methods=["GET"])
def health_check():
    return jsonify({"status": "ok"})

# ----------------------------
# 工具函数
# ----------------------------
def allowed_file(filename: str) -> bool:
    return filename.endswith(".xlsx") or filename.endswith(".xls")

# ----------------------------
# 上传和下载示例（保持你原来的逻辑）
# ----------------------------
@app.route("/api/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"success": False, "message": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "" or not allowed_file(file.filename):
        return jsonify({"success": False, "message": "Invalid file"}), 400

    upload_folder = os.path.join(os.getcwd(), "uploads")
    os.makedirs(upload_folder, exist_ok=True)
    filepath = os.path.join(upload_folder, file.filename)
    file.save(filepath)
    return jsonify({"success": True, "filename": file.filename})

import os
from flask import send_file, jsonify
from urllib.parse import quote  # 用于中文文件名编码

@app.route("/api/download", methods=["GET"])
def download_file():
    """
    下载汇总结果文件，适配 Render / Gunicorn，支持中文文件名
    """
    # 拼接输出文件夹路径（绝对路径）
    output_folder = os.path.join(os.getcwd(), "output")
    os.makedirs(output_folder, exist_ok=True)

    # 输出文件路径
    output_file = os.path.join(output_folder, "汇总结果.xlsx")

    # 如果文件不存在，返回 404
    if not os.path.exists(output_file):
        return jsonify({"success": False, "message": "File not found"}), 404

    # 中文文件名编码，确保在 Linux / 浏览器都能正确下载
    download_name = "汇总结果.xlsx"
    try:
        quoted_name = quote(download_name)
    except:
        quoted_name = download_name

    # 返回文件
    return send_file(
        output_file,
        as_attachment=True,
        download_name=quoted_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------
# 不要写 app.run()
# ----------------------------
# Render 会用 gunicorn 启动你的 Flask app


