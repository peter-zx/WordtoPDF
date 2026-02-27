@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo   WordtoPDF 安装程序
echo ========================================
echo.

REM 检查Python是否安装
echo [1/5] 检查Python环境...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.8或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [成功] Python版本: %PYTHON_VERSION%

REM 检查pip
echo [2/5] 检查pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] pip未安装，请重新安装Python
    pause
    exit /b 1
)
echo [成功] pip已安装

REM 安装依赖
echo [3/5] 安装依赖包...
echo 这可能需要几分钟，请耐心等待...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)
echo [成功] 依赖安装完成

REM 创建桌面快捷方式
echo [4/5] 创建桌面快捷方式...
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\WordtoPDF.lnk"
set "TARGET_PATH=%~dp0run.bat"
set "ICON_PATH=%~dp0assets\images\icon.ico"

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%TARGET_PATH%'; $s.WorkingDirectory = '%~dp0'; $s.Description = 'WordtoPDF - Word转PDF工具'; if (Test-Path '%ICON_PATH%') { $s.IconLocation = '%ICON_PATH%' }; $s.Save()"

if exist "%SHORTCUT_PATH%" (
    echo [成功] 桌面快捷方式已创建
) else (
    echo [警告] 快捷方式创建失败，您可以手动运行 run.bat
)

REM 完成提示
echo.
echo ========================================
echo   安装完成！
echo ========================================
echo.
echo 启动方式:
echo   1. 双击桌面快捷方式 "WordtoPDF"
echo   2. 或双击运行 run.bat
echo.
echo 应用将在浏览器中打开: http://127.0.0.1:8501
echo.

set /p LAUNCH="是否立即启动应用? (Y/N): "
if /i "%LAUNCH%"=="Y" (
    echo.
    echo 正在启动应用...
    start "" "%TARGET_PATH%"
    echo 应用已启动，请查看浏览器窗口
)

echo.
echo 感谢使用 WordtoPDF！
pause
