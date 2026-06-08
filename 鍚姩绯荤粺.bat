@echo off
chcp 65001 > nul
echo.
echo ============================================
echo   外贸单据自动生成系统 v2.0
echo ============================================
echo.

REM 检查Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 安装依赖
echo 正在检查并安装依赖...
pip install flask pandas openpyxl xlrd num2words -q

echo.
echo 系统启动中...
echo.
echo 请在浏览器打开：http://localhost:5000
echo.
echo （关闭此窗口即可停止系统）
echo.

python app.py

pause
