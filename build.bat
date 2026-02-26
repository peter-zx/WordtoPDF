@echo off
chcp 65001 >nul
echo ========================================
echo   WordtoPDF 应用打包
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

echo [1/4] 安装打包工具...
pip install pyinstaller -q

echo [2/4] 安装依赖包...
pip install -r requirements.txt -q

echo [3/4] 清理旧文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo [4/4] 开始打包...
pyinstaller --name "WordtoPDF" ^
    --onefile ^
    --windowed ^
    --icon=assets/icon.ico ^
    --add-data "pages;pages" ^
    --add-data "components;components" ^
    --add-data "services;services" ^
    --add-data "assets;assets" ^
    --add-data "app.py;." ^
    --add-data "config.py;." ^
    --hidden-import=streamlit ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=python-docx ^
    --hidden-import=beautifulsoup4 ^
    --hidden-import=striprtf ^
    --hidden-import=win32com ^
    --hidden-import=comtypes ^
    launcher.py

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo   打包完成！
    echo ========================================
    echo.
    echo 可执行文件位置: dist\WordtoPDF.exe
    echo.
    echo 您可以将 dist\WordtoPDF.exe 发送给其他用户使用
    echo.
) else (
    echo.
    echo [错误] 打包失败，请检查上方错误信息
    echo.
)

pause
