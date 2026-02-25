"""
批量DOCX转PDF服务
"""

import os
import sys
import time
from typing import List, Tuple, Dict
from pathlib import Path


class BatchPDFService:
    """批量DOCX转PDF服务"""

    @staticmethod
    def scan_folder_structure(folder_path: str) -> Dict:
        """
        扫描文件夹结构,识别所有DOCX文件

        Args:
            folder_path: 文件夹路径

        Returns:
            文件夹结构字典
        """
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
                    # 递归扫描子文件夹
                    child_structure = BatchPDFService.scan_folder_structure(item_path)
                    structure["children"].append(child_structure)

                elif item.lower().endswith('.docx'):
                    # 添加DOCX文件
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
        """
        从文件夹结构中获取所有DOCX文件

        Args:
            structure: 文件夹结构

        Returns:
            DOCX文件列表
        """
        files = []

        # 添加当前文件夹的DOCX文件
        files.extend(structure.get("docx_files", []))

        # 递归添加子文件夹的DOCX文件
        for child in structure.get("children", []):
            files.extend(BatchPDFService.get_all_docx_files(child))

        return files

    @staticmethod
    def convert_docx_to_pdf(input_path: str, output_path: str) -> Tuple[bool, str]:
        """
        将DOCX文件转换为PDF

        Args:
            input_path: 输入DOCX文件路径
            output_path: 输出PDF文件路径

        Returns:
            (成功标志, 错误信息)
        """
        word_app = None
        try:
            import win32com.client

            # 创建Word应用实例
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False

            # 打开DOCX文件
            doc = word_app.Documents.Open(input_path, ReadOnly=True)

            # 保存为PDF格式
            doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF format

            # 关闭文档
            doc.Close()

            # 关闭Word应用
            word_app.Quit()

            return True, ""

        except Exception as e:
            # 如果出错，确保关闭Word应用
            if word_app is not None:
                try:
                    word_app.Quit()
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
        批量转换DOCX文件,保持文件夹结构

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

        for index, file_info in enumerate(all_files, 1):
            try:
                # 计算相对路径
                relative_path = file_info["relative_path"]

                # 保持文件夹结构
                relative_dir = os.path.dirname(relative_path)
                output_dir = os.path.join(output_base_dir, relative_dir)

                # 确保输出目录存在
                os.makedirs(output_dir, exist_ok=True)

                # 生成输出文件名
                input_name = Path(file_info["name"]).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")

                # 转换文件
                success, error = BatchPDFService.convert_docx_to_pdf(
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

                # 每次转换后稍作延迟，避免Word进程冲突
                time.sleep(0.5)

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
        批量将DOCX文件转换为PDF

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

        for index, input_file in enumerate(input_files, 1):
            try:
                # 生成输出文件名
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")

                # 转换文件
                success, error = BatchPDFService.convert_docx_to_pdf(
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

                # 每次转换后稍作延迟，避免Word进程冲突
                time.sleep(0.5)

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
        """
        获取转换结果摘要

        Args:
            results: 转换结果列表

        Returns:
            摘要信息
        """
        total = len(results)
        success_count = sum(1 for r in results if r["success"])
        failed_count = total - success_count

        return {
            "total": total,
            "success": success_count,
            "failed": failed_count,
            "success_rate": f"{(success_count/total*100):.1f}%" if total > 0 else "0%"
        }
