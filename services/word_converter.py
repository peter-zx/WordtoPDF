"""
Word转PDF工具 - 单文件转换（修复版本）
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

        # 最终判断：检查输出文件是否存在
        # 即使subprocess返回错误，如果文件存在就认为成功
        if os.path.exists(output_path):
            # 检查文件大小，确保不是空文件
            if os.path.getsize(output_path) > 0:
                return True, ""
            else:
                return False, "输出文件为空"

        # 文件不存在，返回错误信息
        if result.returncode == 0 and "SUCCESS" in result.stdout:
            return False, "输出文件未生成"
        else:
            error_msg = result.stderr or result.stdout or "未知错误"
            return False, error_msg

    except subprocess.TimeoutExpired:
        # 超时，但检查文件是否已生成
        try:
            os.remove(script_path)
        except:
            pass

        # 即使超时，如果文件存在且有效，也算成功
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return True, ""
        return False, "转换超时"

    except Exception as e:
        # 清理脚本
        try:
            os.remove(script_path)
        except:
            pass

        # 即使异常，如果文件存在且有效，也算成功
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return True, ""
        return False, str(e)
