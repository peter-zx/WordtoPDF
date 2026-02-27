# 服务器部署脚本 - WordtoPDF
# 使用方法：ssh ubuntu@122.51.231.239 'bash -s' < deploy_to_server.sh

echo "=========================================="
echo "WordtoPDF 服务器部署脚本"
echo "=========================================="

# 1. 停止旧服务
echo ""
echo "🛑 步骤 1: 停止旧的 Streamlit 服务..."
pkill -f "streamlit run.*app.py" || true
pkill -f "streamlit run.*批量 PDF 转换" || true
sleep 2

# 检查是否停止成功
if pgrep -f "streamlit run.*app.py" > /dev/null; then
    echo "❌ 无法停止旧服务，请手动检查"
    exit 1
else
    echo "✅ 旧服务已停止"
fi

# 2. 创建应用目录
echo ""
echo "📁 步骤 2: 创建应用目录..."
mkdir -p ~/apps/wordtopdf
cd ~/apps/wordtopdf

# 3. 创建虚拟环境 (如果不存在)
echo ""
echo "🐍 步骤 3: 配置 Python 虚拟环境..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "✅ 虚拟环境已创建"
else
    echo "✅ 虚拟环境已存在"
fi

# 激活虚拟环境
source venv/bin/activate

# 4. 安装依赖
echo ""
echo "📦 步骤 4: 安装 Python 依赖..."
if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt --upgrade
    echo "✅ 依赖已安装"
else
    echo "⚠️ 未找到 requirements.txt"
fi

# 5. 创建 systemd 服务文件
echo ""
echo "⚙️ 步骤 5: 配置 systemd 持久化服务..."
sudo tee /etc/systemd/system/wordtopdf.service > /dev/null <<EOF
[Unit]
Description=WordtoPDF Streamlit Application
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/apps/wordtopdf
Environment="PATH=/home/ubuntu/apps/wordtopdf/venv/bin"
ExecStart=/home/ubuntu/apps/wordtopdf/venv/bin/streamlit run app.py --server.port 8501 --server.address 0.0.0.0
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

echo "✅ systemd 服务文件已创建"

# 6. 重载 systemd 并启动服务
echo ""
echo "🚀 步骤 6: 启动服务..."
sudo systemctl daemon-reload
sudo systemctl enable wordtopdf
sudo systemctl restart wordtopdf

# 7. 检查服务状态
echo ""
echo "📊 步骤 7: 检查服务状态..."
sleep 3
sudo systemctl status wordtopdf --no-pager

# 8. 显示访问信息
echo ""
echo "=========================================="
echo "✅ 部署完成!"
echo "=========================================="
echo ""
echo "🌐 访问地址："
echo "  内网：http://localhost:8501"
echo "  外网：http://122.51.231.239:8501"
echo ""
echo "🔧 服务管理命令:"
echo "  查看状态：sudo systemctl status wordtopdf"
echo "  停止服务：sudo systemctl stop wordtopdf"
echo "  重启服务：sudo systemctl restart wordtopdf"
echo "  查看日志：sudo journalctl -u wordtopdf -f"
echo ""
echo "=========================================="
