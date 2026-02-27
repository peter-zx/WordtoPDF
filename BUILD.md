# WordtoPDF Application Build Guide

## Build Steps

### 1. Prepare Environment
- Install Python 3.8 or higher
- Install Microsoft Word (required for PDF conversion)

### 2. Install Dependencies
```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 3. Build Application
Double-click `build.bat` or run:
```bash
build.bat
```

### 4. Get Distribution
After building, you'll find:
```
dist/WordtoPDF/          - Application folder
dist/WordtoPDF.zip       - Compressed package (recommended for distribution)
```

## User Instructions

### Send to Users
Send the `dist/WordtoPDF.zip` file to users

### Installation
1. Extract `WordtoPDF.zip`
2. Open the extracted `WordtoPDF` folder
3. Double-click `WordtoPDF.exe` to launch

### Usage
1. Application will start automatically
2. Browser opens at http://127.0.0.1:8501
3. Start using the application

## System Requirements

- Windows 10/11 64-bit
- Microsoft Word (for PDF conversion feature)
- 4GB RAM or higher (8GB recommended)
- 500MB free disk space

## Technical Details

### Build Mode
- **Mode**: Directory mode (onedir)
- **Reason**: Reduces memory usage during execution
- **Entry Point**: launcher.py

### Package Contents
```
WordtoPDF/
├── WordtoPDF.exe      (20MB - main executable)
├── _internal/         (Python runtime and dependencies)
└── pages/             (Application pages)
```

### Distribution Format
- **Folder**: dist/WordtoPDF/ (uncompressed)
- **ZIP**: dist/WordtoPDF.zip (120MB compressed)

## Troubleshooting

### Build Issues
- Ensure Python version >= 3.8
- Check all dependencies are installed
- Verify Windows version (10/11 64-bit only)

### Runtime Issues
- Check Microsoft Word is installed
- Verify sufficient RAM (4GB minimum)
- Check antivirus software (may block execution)
- Verify Windows Defender settings

### Performance Tips
- Close other applications before running
- Use SSD for better performance
- Increase virtual memory if needed

## Advanced Options

### Custom Build
Edit `build.bat` to modify:
- `--onedir`: Use directory mode
- `--windowed`: Hide console window
- `--add-data`: Add additional files
- `--hidden-import`: Include extra modules

### Single File Mode
To build as single exe (larger, more memory):
```bash
pyinstaller --name "WordtoPDF" --onefile ...
```
**Warning**: May cause memory issues on systems with limited RAM

## Version History

### v1.0 (2026-02-27)
- Initial release
- Directory mode packaging
- Reduced memory usage
- 20MB executable size
