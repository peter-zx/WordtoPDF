# 腾讯云服务器部署指南

## 📋 服务器信息
- **IP 地址**: 122.51.231.239
- **用户名**: ubuntu
- **密码**: Aa112233#
- **SSH 端口**: 22
- **应用路径**: /home/ubuntu/apps/wordtopdf
- **访问端口**: 8501

## 🚀 快速部署步骤

### 方法一：使用部署脚本 (推荐)

1. **SSH 登录服务器**
```bash
ssh ubuntu@122.51.231.239
# 输入密码：Aa112233#
```

2. **停止旧服务**
```bash
sudo systemctl stop wordtopdf
pkill -f streamlit
```

3. **拉取最新代码**
```bash
cd ~/apps/wordtopdf
git pull origin wenjianjia
```

4. **运行部署脚本**
```bash
chmod +x deploy_to_server.sh
./deploy_to_server.sh
```

### 方法二：手动部署

如果自动脚本有问题，可以手动执行以下步骤:

#### 1. 停止旧服务
```bash
# 查找进程
ps aux | grep streamlit

# 停止 systemd 服务
sudo systemctl stop wordtopdf

# 强制停止所有 streamlit 进程
pkill -9 -f streamlit
```

#### 2. 更新代码
```bash
cd ~/apps/wordtopdf
git pull origin wenjianjia
```

#### 3. 激活虚拟环境并安装依赖
```bash
cd ~/apps/wordtopdf
source venv/bin/activate
pip install -r requirements.txt --upgrade
```

#### 4. 配置 systemd 服务
```bash
sudo nano /etc/systemd/system/wordtopdf.service
```

粘贴以下内容:
```ini
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
```

保存退出 (Ctrl+O, Enter, Ctrl+X)

#### 5. 启动服务
```bash
sudo systemctl daemon-reload
sudo systemctl enable wordtopdf
sudo systemctl restart wordtopdf
```

#### 6. 检查状态
```bash
sudo systemctl status wordtopdf
```

## 🔧 服务管理命令

### 查看服务状态
```bash
sudo systemctl status wordtopdf
```

### 停止服务
```bash
sudo systemctl stop wordtopdf
```

### 启动服务
```bash
sudo systemctl start wordtopdf
```

### 重启服务
```bash
sudo systemctl restart wordtopdf
```

### 查看日志
```bash
# 实时查看日志
sudo journalctl -u wordtopdf -f

# 查看最近 100 行
sudo journalctl -u wordtopdf -n 100
```

### 禁用开机自启
```bash
sudo systemctl disable wordtopdf
```

## 🌐 访问应用

部署完成后，可以通过以下方式访问:

- **外网访问**: http://122.51.231.239:8501
- **本地访问**: http://localhost:8501

## 🔒 安全配置 (可选)

### 配置腾讯云安全组

如果无法访问 8501 端口，需要在腾讯云控制台开放端口:

1. 登录腾讯云控制台
2. 进入安全组配置
3. 添加入站规则:
   - 端口：8501
   - 协议：TCP
   - 来源：0.0.0.0/0 (或指定 IP)

### 修改访问端口

如果要修改默认端口，编辑 systemd 服务文件:
```bash
sudo nano /etc/systemd/system/wordtopdf.service
```

修改 ExecStart 行中的端口号:
```
--server.port 8501  # 改成其他端口，如 8080
```

然后重启服务:
```bash
sudo systemctl daemon-reload
sudo systemctl restart wordtopdf
```

## 📊 环境隔离说明

本部署使用 Python 虚拟环境实现环境隔离:

- **虚拟环境路径**: `/home/ubuntu/apps/wordtopdf/venv`
- **Python 版本**: 系统默认的 Python 3
- **依赖包位置**: `/home/ubuntu/apps/wordtopdf/venv/lib/pythonX.X/site-packages`

每个应用都有独立的虚拟环境，互不影响。后续安装其他应用时:
1. 创建新的目录 (如 `~/apps/other-app`)
2. 创建新的虚拟环境
3. 安装各自的依赖

## ⚠️ 常见问题

### 问题 1: 服务启动失败
```bash
# 查看详细错误
sudo journalctl -u wordtopdf -n 50

# 检查端口是否被占用
sudo netstat -tlnp | grep 8501

# 检查权限
ls -la /home/ubuntu/apps/wordtopdf
```

### 问题 2: 无法访问网页
```bash
# 检查防火墙
sudo ufw status

# 检查安全组 (腾讯云控制台)
# 确保 8501 端口已开放
```

### 问题 3: 虚拟环境问题
```bash
# 删除虚拟环境重新创建
cd ~/apps/wordtopdf
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 问题 4: Git 拉取失败
```bash
# 如果遇到冲突，可以强制重置
cd ~/apps/wordtopdf
git fetch origin wenjianjia
git reset --hard origin/wenjianjia
```

## 📝 下次更新流程

当需要更新到新版本时:

```bash
# 1. SSH 登录
ssh ubuntu@122.51.231.239

# 2. 停止服务
sudo systemctl stop wordtopdf

# 3. 拉取最新代码
cd ~/apps/wordtopdf
git pull origin wenjianjia

# 4. 重启服务
sudo systemctl start wordtopdf

# 5. 检查状态
sudo systemctl status wordtopdf
```

## ✅ 部署完成检查清单

- [ ] 旧服务已停止
- [ ] 新代码已拉取
- [ ] 虚拟环境已配置
- [ ] 依赖已安装
- [ ] systemd 服务已配置
- [ ] 服务已启动
- [ ] 可以通过 http://122.51.231.239:8501 访问
- [ ] 批量 PDF 转换功能正常
- [ ] 文件夹生成功能正常

---

**技术支持**: 查看日志文件或联系管理员
