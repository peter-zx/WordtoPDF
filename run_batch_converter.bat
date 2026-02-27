@echo off
chcp 65001 >nul
echo ========================================
echo   批量 PDF 转换工具 - 独立版本
echo ========================================
echo.
echo 正在启动服务器...
echo.

cd /d "%~dp0"
python -m streamlit run apps\批量 PDF 转换.py --server.port 8505 --server.address 0.0.0.0

pause
