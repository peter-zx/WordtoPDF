"""
批量DOCX转PDF服务 - 最简单串行版本
一个接一个转换，不并发，最简单最稳定
"""

import os
import sys
import time
import subprocess
from typing import List, Tuple, Dict
from pathlib import Path


class BatchPDFService:
    """批量DOCX转PDF服务 - 最简单串行版本"""

    @staticmethod
    def scan_folder_structure(folder_path: str) -> Dict:
        """扫描文件夹结构,识别所有DOCX文件"""
        structure = {
            "path": folder_path,
            "name": os.path.basename(folder_path),
            "type": "folder",
            "children": [],
            "docx_files": []
        }

        try:
            items = sorted(os.listdir(folder_path))

            for item in items:
                item_path = os.path.join(folder_path, item)

                if os.path.isdir(item_path):
                    child_structure = BatchPDFService.scan_folder_structure(item_path)
                    structure["children"].append(child_structure)

                elif item.lower().endswith('.docx'):
                    structure["docx_files"].append({
                        "name": item,
                        "path": item_path,
                        "relative_path": os.path.relpath(item_path, folder_path)
                    })

        except Exception as e:
            print(f"扫描文件夹失败: {str(e)}")

        return structure

    @staticmethod
    def get_all_docx_files(structure: Dict) -> List[Dict]:
        """从文件夹结构中获取所有DOCX文件"""
        files = []
        files.extend(structure.get("docx_files", []))
        for child in structure.get("children", []):
            files.extend(BatchPDFService.get_all_docx_files(child))
        return files

    @staticmethod
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
    
    # 打开文档
    doc = word.Documents.Open(r"{input_path}", ReadOnly=True, Visible=False)
    
    # 保存为PDF
    doc.SaveAs(r"{output_path}", FileFormat=17)
    
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

        # 写入临时脚本
        script_path = os.path.join(os.path.dirname(output_path), "convert_script.py")
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

    @staticmethod
    def batch_convert_with_structure(
        structure: Dict,
        output_base_dir: str,
        progress_callback=None
    ) -> List[dict]:
        """
        批量转换DOCX文件,保持文件夹结构（最简单串行版本）

        Args:
            structure: 文件夹结构
            output_base_dir: 输出基础目录
            progress_callback: 进度回调函数

        Returns:
            转换结果列表
        """
        # 获取所有DOCX文件
        all_files = BatchPDFService.get_all_docx_files(structure)
        total = len(all_files)

        if total == 0:
            return []

        results = []

        # 获取顶层文件夹名称
        top_folder_name = structure["name"]

        # 创建顶层文件夹
        top_output_dir = os.path.join(output_base_dir, top_folder_name)
        os.makedirs(top_output_dir, exist_ok=True)

        # 逐个处理文件（串行，一个接一个）
        for index, file_info in enumerate(all_files, 1):
            try:
                # 计算相对路径
                relative_path = file_info["relative_path"]

                # 保持文件夹结构，添加顶层文件夹
                relative_dir = os.path.dirname(relative_path)
                output_dir = os.path.join(top_output_dir, relative_dir)

                # 确保输出目录存在
                os.makedirs(output_dir, exist_ok=True)

                # 生成输出文件名
                input_name = Path(file_info["name"]).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")

                # 转换文件（独立进程）
                success, error = BatchPDFService.convert_single_file(
                    file_info["path"],
                    output_file
                )

                result = {
                    "input_file": file_info["path"],
                    "output_file": output_file if success else None,
                    "relative_path": relative_path,
                    "success": success,
                    "error": error if not success else None
                }
                results.append(result)

                # 调用进度回调
                if progress_callback:
                    progress_callback(index, total, result)

                # 文件间延迟，确保进程完全退出
                time.sleep(2)

            except Exception as e:
                result = {
                    "input_file": file_info["path"],
                    "output_file": None,
                    "relative_path": file_info.get("relative_path", ""),
                    "success": False,
                    "error": str(e)
                }
                results.append(result)

                if progress_callback:
                    progress_callback(index, total, result)

        return results

    @staticmethod
    def batch_convert_docx_to_pdf(
        input_files: List[str],
        output_dir: str,
        progress_callback=None
    ) -> List[dict]:
        """
        批量将DOCX文件转换为PDF（最简单串行版本）

        Args:
            input_files: 输入DOCX文件路径列表
            output_dir: 输出目录
            progress_callback: 进度回调函数

        Returns:
            转换结果列表
        """
        results = []
        total = len(input_files)

        if total == 0:
            return []

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 逐个处理文件（串行）
        for index, input_file in enumerate(input_files, 1):
            try:
                # 生成输出文件名
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")

                # 转换文件（独立进程）
                success, error = BatchPDFService.convert_single_file(
                    input_file,
                    output_file
                )

                result = {
                    "input_file": input_file,
                    "output_file": output_file if success else None,
                    "success": success,
                    "error": error if not success else None
                }
                results.append(result)

                # 调用进度回调
                if progress_callback:
                    progress_callback(index, total, result)

                # 文件间延迟
                time.sleep(2)

            except Exception as e:
                result = {
                    "input_file": input_file,
                    "output_file": None,
                    "success": False,
                    "error": str(e)
                }
                results.append(result)

                if progress_callback:
                    progress_callback(index, total, result)

        return results

    @staticmethod
    def get_conversion_summary(results: List[dict]) -> dict:
        """获取转换结果摘要"""
        total = len(results)
        success_count = sum(1 for r in results if r["success"])
        failed_count = total - success_count

        return {
            "total": total,
            "success": success_count,
            "failed": failed_count,
            "success_rate": f"{(success_count/total*100):.1f}%" if total > 0 else "0%"
        }
