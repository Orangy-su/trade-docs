@echo off
chcp 65001 > nul
echo.
echo ================================================
echo   外贸单据自动生成系统 v2.0  (主数据固化版)
echo ================================================
echo.
python --version > nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载：https://www.python.org/downloads/
    pause & exit /b 1
)
echo 安装依赖...
pip install -r requirements.txt -q
echo.
echo 启动成功！请在浏览器打开：
echo   http://localhost:5000
echo.
echo 主数据手册上传一次后会自动保存，下次无需重新上传。
echo.
echo （关闭此窗口停止服务）
echo.
python app_v2.py
pause
