@echo off
REM ============================================
REM 开票汇总助手一键启动（自动打开浏览器+自动保存并打开Excel）
REM ============================================

REM 设置汇总结果存放路径
set RESULT_DIR=%USERPROFILE%\Desktop\汇总结果
if not exist "%RESULT_DIR%" (
    mkdir "%RESULT_DIR%"
)

REM -----------------------------
REM 启动后端
REM -----------------------------
cd backend
python -m pip install --upgrade pip
python -m pip install pandas --prefer-binary
python -m pip install openpyxl flask flask-cors
python -m pip install -r requirements.txt

echo 启动后端 Flask 服务...
start cmd /k "set RESULT_DIR=%RESULT_DIR% && python app.py"

REM -----------------------------
REM 启动前端
REM -----------------------------
cd ..
cd frontend
npm install
start cmd /k "npm run dev"

REM -----------------------------
REM 打开浏览器
REM -----------------------------
timeout /t 5 /nobreak >nul
start http://localhost:5173

echo 开票汇总助手已启动成功！
pause