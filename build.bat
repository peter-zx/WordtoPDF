@echo off
echo ========================================
echo   WordtoPDF Application Build
echo ========================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.8 or higher.
    pause
    exit /b 1
)

echo [1/4] Installing PyInstaller...
pip install pyinstaller -q

echo [2/4] Installing dependencies...
pip install -r requirements.txt -q

echo [3/4] Cleaning old files...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo [4/4] Building application...
pyinstaller --name "WordtoPDF" ^
    --onedir ^
    --windowed ^
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
    echo   Build Complete!
    echo ========================================
    echo.
    echo Application folder: dist\WordtoPDF\
    echo.
    echo You can send the entire dist\WordtoPDF folder to other users.
    echo.
) else (
    echo.
    echo [ERROR] Build failed. Please check error messages above.
    echo.
)

pause
