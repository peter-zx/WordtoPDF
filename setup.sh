#!/bin/bash

echo "========================================"
echo "文件批量处理工具箱 - Linux/Mac 启动脚本"
echo "========================================"
echo ""
echo "正在启动..."
echo ""

# 激活虚拟环境
source venv/bin/activate

if [ $? -ne 0 ]; then
    echo "[错误] 无法激活虚拟环境!"
    echo "请确保 venv 文件夹存在"
    exit 1
fi

echo "[成功] 虚拟环境已激活"
echo ""
echo "浏览器将自动打开"
echo "如果没有打开请手动访问: http://localhost:8501"
echo ""
echo "按Ctrl+C可以停止程序"
echo "========================================"
echo ""

streamlit run app.py
