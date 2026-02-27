@echo off
chcp 65001 >nul
echo ========================================
echo   WordtoPDF 代码推送脚本
echo ========================================
echo.

cd /d "%~dp0"

echo [1/3] 检查Git状态...
git status

echo.
echo [2/3] 查看未提交的更改...
git diff --name-only

echo.
echo [3/3] 请填写版本说明（留空取消）：
echo ========================================
set /p MESSAGE="请输入版本说明: "

if "%MESSAGE%"=="" (
    echo.
    echo [取消] 未输入版本说明，推送已取消
    pause
    exit /b 0
)

echo.
echo 正在提交...
echo ========================================

git add .
git commit -m "%MESSAGE%

🤖 Generated with CodeMate"

if %errorlevel% neq 0 (
    echo.
    echo [错误] 提交失败
    pause
    exit /b 1
)

echo.
echo 正在推送到远程仓库...
echo ========================================

git push origin wenjianjia

if %errorlevel% neq 0 (
    echo.
    echo [错误] 推送失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo   推送成功！
echo ========================================
echo.
pause
