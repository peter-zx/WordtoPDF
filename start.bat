@echo off
REM WordtoPDF 应用启动脚本
REM 用于打包后的exe文件

echo 正在启动 WordtoPDF...
echo.
echo 应用将在浏览器中打开: http://127.0.0.1:8501
echo 请保持此窗口运行
echo.
echo 按 Ctrl+C 可停止应用
echo ========================================

cd /d "%~dp0"
WordtoPDF.exe

pause
