#!/bin/bash
echo ""
echo "============================================"
echo "  外贸单据自动生成系统 v2.0"
echo "============================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "错误：未检测到 Python3，请先安装"
    exit 1
fi

# 安装依赖
echo "检查依赖..."
pip3 install flask pandas openpyxl xlrd num2words -q

echo ""
echo "系统启动中..."
echo "请在浏览器打开：http://localhost:5000"
echo ""
echo "（按 Ctrl+C 停止系统）"
echo ""

python3 app.py
