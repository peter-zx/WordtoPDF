@echo off
chcp 65001 >nul
echo ==========================================
echo WordtoPDF 腾讯云服务器部署工具
echo ==========================================
echo.
echo 服务器：122.51.231.239
echo 用户：ubuntu  
echo 路径：/home/ubuntu/apps/wordtopdf
echo.
echo 正在连接服务器...
echo.

REM 使用 Plink (PuTTY 组件) 或 OpenSSH 执行远程命令
REM 方法 1: 如果有 plink
plink -ssh ubuntu@122.51.231.239 -pw "Aa112233#" -P 22 ^
"cd ~/apps/wordtopdf && ^
sudo systemctl stop wordtopdf && ^
pkill -f streamlit || true && ^
git pull origin wenjianjia && ^
source venv/bin/activate && ^
pip install -r requirements.txt --upgrade && ^
sudo systemctl daemon-reload && ^
sudo systemctl enable wordtopdf && ^
sudo systemctl restart wordtopdf && ^
sleep 3 && ^
sudo systemctl status wordtopdf --no-pager"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==========================================
    echo ✅ 部署完成!
    echo ==========================================
    echo.
    echo 访问地址：http://122.51.231.239:8501
    echo.
) else (
    echo.
    echo ==========================================
    echo ❌ 部署失败，请手动执行
    echo ==========================================
    echo.
    echo 请按以下步骤手动部署:
    echo.
    echo 1. 打开 PowerShell 或 CMD
    echo 2. 执行：ssh ubuntu@122.51.231.239
    echo 3. 输入密码：Aa112233#
    echo 4. 依次执行以下命令:
    echo.
    echo    cd ~/apps/wordtopdf
    echo    sudo systemctl stop wordtopdf
    echo    pkill -f streamlit
    echo    git pull origin wenjianjia
    echo    source venv/bin/activate
    echo    pip install -r requirements.txt --upgrade
    echo    sudo systemctl daemon-reload
    echo    sudo systemctl enable wordtopdf
    echo    sudo systemctl restart wordtopdf
    echo    sudo systemctl status wordtopdf
    echo.
)

pause
