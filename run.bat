@echo off
echo ========================================
echo 文件批量处理工具箱
echo ========================================
echo.
echo 正在启动...
echo.

REM 激活虚拟环境
call venv\Scripts\activate.bat

if errorlevel 1 (
    echo [错误] 无法激活虚拟环境!
    echo 请确保 venv 文件夹存在
    pause
    exit /b 1
)

echo [成功] 虚拟环境已激活
echo.
echo 浏览器将自动打开
echo 如果没有打开请手动访问: http://localhost:8501
echo.
echo 按Ctrl+C可以停止程序
echo ========================================
echo.

streamlit run app.py

pause
