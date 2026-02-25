@echo off
echo ========================================
echo 文件批量处理工具箱 - 完整安装脚本
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
echo 正在安装Python依赖包...

call venv\Scripts\activate.bat
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

if errorlevel 1 (
    echo [错误] 安装依赖失败!
    pause
    exit /b 1
)

echo.
echo ========================================
echo 正在检查系统依赖...
echo ========================================

REM 检查pandoc
echo.
echo 检查 Pandoc...
python -c "import pypandoc; print('[成功] Pandoc已安装:', pypandoc.get_pandoc_path())" 2>nul
if errorlevel 1 (
    echo [警告] Pandoc未安装，正在尝试自动安装...
    echo 注意: 如果自动安装失败，请手动下载安装 Pandoc
    echo 下载地址: https://pandoc.org/installing.html
    echo.
    echo 尝试使用 winget 安装...
    winget install --id JohnMacFarlane.Pandoc -e --accept-source-agreements --accept-package-agreements
    if errorlevel 1 (
        echo [警告] winget安装失败，请手动安装Pandoc
    ) else (
        echo [成功] Pandoc安装成功
    )
)

REM 检查pywin32
echo.
echo 检查 pywin32...
python -c "import win32com.client; print('[成功] pywin32已安装')" 2>nul
if errorlevel 1 (
    echo [警告] pywin32未安装，正在安装...
    pip install pywin32
    if errorlevel 1 (
        echo [错误] pywin32安装失败
    ) else (
        echo [成功] pywin32安装成功
    )
)

echo.
echo ========================================
echo [成功] 安装完成!
echo ========================================
echo.
echo 运行程序: 双击 run.bat 文件
echo 或执行命令: streamlit run app.py
echo.
pause
