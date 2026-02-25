"""
Word转PDF工具 - 单文件转换
"""

import os
import sys
import subprocess
from typing import Tuple


def convert_single_file(input_path: str, output_path: str) -> Tuple[bool, str]:
    """
    转换单个文件 - 使用独立Word进程

    Args:
        input_path: 输入DOCX文件路径
        output_path: 输出PDF文件路径

    Returns:
        (成功标志, 错误信息)
    """
    # 创建一个Python脚本文件来执行转换
    script_content = f'''
import sys
import os
import win32com.client
import time

try:
    # 创建Word实例
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    # 打开文档（绝对路径）
    input_file = r"{input_path}"
    output_file = r"{output_path}"
    
    doc = word.Documents.Open(input_file, ReadOnly=True, Visible=False)
    
    # 保存为PDF
    doc.SaveAs(output_file, FileFormat=17)
    
    # 关闭文档
    doc.Close()
    
    # 关闭Word
    word.Quit()
    
    # 等待Word完全退出
    time.sleep(1)
    
    print("SUCCESS")
    sys.exit(0)
    
except Exception as e:
    print(f"ERROR: {{str(e)}}")
    # 尝试关闭Word
    try:
        if 'word' in locals():
            word.Quit()
    except:
        pass
    sys.exit(1)
'''

    # 写入临时脚本（放在输出目录下）
    script_dir = os.path.dirname(output_path)
    if not os.path.exists(script_dir):
        os.makedirs(script_dir, exist_ok=True)

    script_path = os.path.join(script_dir, "convert_script.py")
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(script_content)

    try:
        # 执行脚本
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        # 清理临时脚本
        try:
            os.remove(script_path)
        except:
            pass

        if result.returncode == 0 and "SUCCESS" in result.stdout:
            if os.path.exists(output_path):
                return True, ""
            else:
                return False, "输出文件未生成"
        else:
            error_msg = result.stderr or result.stdout or "未知错误"
            return False, error_msg

    except subprocess.TimeoutExpired:
        # 超时，清理脚本
        try:
            os.remove(script_path)
        except:
            pass
        return False, "转换超时"

    except Exception as e:
        # 清理脚本
        try:
            os.remove(script_path)
        except:
            pass
        return False, str(e)
