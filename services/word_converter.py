"""
Word转PDF工具 - 跨平台支持
- Windows: 使用Microsoft Word (pywin32)
- Linux: 使用LibreOffice (soffice命令)
"""

import os
import sys
import subprocess
import platform
from typing import Tuple


def is_windows() -> bool:
    """检查是否为Windows系统"""
    return platform.system() == 'Windows'


def is_libreoffice_available() -> bool:
    """检查LibreOffice是否可用"""
    try:
        result = subprocess.run(
            ['soffice', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.returncode == 0
    except:
        return False


def convert_with_libreoffice(input_path: str, output_dir: str) -> Tuple[bool, str]:
    """
    使用LibreOffice转换文档为PDF
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录
    
    Returns:
        (成功标志, 错误信息)
    """
    try:
        # LibreOffice命令
        cmd = [
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            input_path
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        # 检查输出文件
        input_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(output_dir, f"{input_name}.pdf")
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return True, ""
        else:
            return False, result.stderr or "转换失败，输出文件不存在"
            
    except subprocess.TimeoutExpired:
        return False, "转换超时"
    except Exception as e:
        return False, str(e)


def convert_with_word(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    使用Microsoft Word转换文档为PDF (Windows专用)
    
    Args:
        input_path: 输入文件路径
        output_path: 输出PDF路径
    
    Returns:
        (成功标志, 错误信息)
    """
    # 创建转换脚本
    script_content = f'''
import sys
import os
import win32com.client
import time

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    input_file = r"{input_path}"
    output_file = r"{output_path}"

    doc = word.Documents.Open(input_file, ReadOnly=True, Visible=False)
    doc.SaveAs(output_file, FileFormat=17)
    doc.Close()
    word.Quit()
    time.sleep(1)

    print("SUCCESS")
    sys.exit(0)

except Exception as e:
    print(f"ERROR: {{str(e)}}")
    try:
        if 'word' in locals():
            word.Quit()
    except:
        pass
    sys.exit(1)
'''

    # 写入临时脚本
    script_dir = os.path.dirname(output_path)
    if not os.path.exists(script_dir):
        os.makedirs(script_dir, exist_ok=True)

    script_path = os.path.join(script_dir, "convert_script.py")
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(script_content)

    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW if is_windows() else 0
        )

        # 清理临时脚本
        try:
            os.remove(script_path)
        except:
            pass

        # 检查输出文件
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return True, ""
        else:
            return False, result.stderr or "转换失败"

    except subprocess.TimeoutExpired:
        try:
            os.remove(script_path)
        except:
            pass
        return False, "转换超时"
    except Exception as e:
        try:
            os.remove(script_path)
        except:
            pass
        return False, str(e)


def convert_single_file(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    转换单个文件为PDF - 自动选择最佳方法
    
    Args:
        input_path: 输入Word文件路径(DOC或DOCX)
        output_path: 输出PDF文件路径
    
    Returns:
        (成功标志, 错误信息)
    """
    # 检查输入文件是否存在
    if not os.path.exists(input_path):
        return False, f"输入文件不存在: {input_path}"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Windows系统优先使用Word
    if is_windows():
        return convert_with_word(input_path, output_path)
    
    # Linux系统使用LibreOffice
    if is_libreoffice_available():
        # LibreOffice输出到目录，需要重命名
        input_name = os.path.splitext(os.path.basename(input_path))[0]
        expected_output = os.path.join(output_dir, f"{input_name}.pdf")
        
        success, error = convert_with_libreoffice(input_path, output_dir)
        
        if success:
            # 如果输出路径不同，重命名
            if expected_output != output_path and os.path.exists(expected_output):
                import shutil
                shutil.move(expected_output, output_path)
            return True, ""
        else:
            return False, error
    
    return False, "无可用的转换工具（需要安装Microsoft Word或LibreOffice）"


def get_converter_info() -> dict:
    """获取当前系统的转换器信息"""
    info = {
        "platform": platform.system(),
        "word_available": False,
        "libreoffice_available": is_libreoffice_available(),
        "recommended_method": None
    }
    
    if is_windows():
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Quit()
            info["word_available"] = True
            info["recommended_method"] = "Microsoft Word"
        except:
            pass
    
    if info["libreoffice_available"]:
        info["recommended_method"] = "LibreOffice"
    
    if not info["recommended_method"]:
        info["recommended_method"] = "无可用转换器"
    
    return info


if __name__ == "__main__":
    # 测试
    print("转换器信息:", get_converter_info())
