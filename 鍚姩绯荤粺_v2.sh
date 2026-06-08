#!/bin/bash
echo ""
echo "================================================"
echo "  外贸单据自动生成系统 v2.0  (主数据固化版)"
echo "================================================"
pip3 install -r requirements.txt -q
echo ""
echo "启动成功！浏览器打开：http://localhost:5000"
echo ""
python3 app_v2.py
