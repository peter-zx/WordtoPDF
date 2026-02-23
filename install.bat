@echo off
echo ========================================
echo 安装文件夹批量生成工具
echo ========================================
echo.
echo 正在创建虚拟环境...

python -m venv venv

if errorlevel 1 (
    echo [错误] 创建虚拟环境失败!
    pause
    exit /b 1
)

echo [成功] 虚拟环境创建完成
echo.
echo 正在安装依赖包...

call venv\Scripts\activate.bat
pip install -r requirements.txt

if errorlevel 1 (
    echo [错误] 安装依赖失败!
    pause
    exit /b 1
)

echo.
echo ========================================
echo [成功] 安装完成!
echo ========================================
echo.
echo 运行程序: 双击 run.bat 文件
echo 或执行命令: streamlit run folder_generator.py
echo.
pause
