"""
Streamlit应用启动器 - 用于PyInstaller打包
"""

import os
import sys
import subprocess

def main():
    # 获取应用所在目录
    app_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 切换到应用目录
    os.chdir(app_dir)
    
    # 启动Streamlit应用
    subprocess.run([
        sys.executable, "-m", "streamlit", "run", "app.py",
        "--server.port", "8501",
        "--server.address", "127.0.0.1"
    ])

if __name__ == "__main__":
    main()
