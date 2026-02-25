"""
批量DOCX转PDF服务 - 使用分离的工具模块
"""

import os
import time
from typing import List, Dict

from services.folder_scanner import scan_folder_structure, get_all_docx_files
from services.word_converter import convert_single_file
from services.path_handler import calculate_output_with_top_folder


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
        return scan_folder_structure(folder_path)

    @staticmethod
    def get_all_docx_files(structure: Dict) -> List[Dict]:
        """
        从文件夹结构中获取所有DOCX文件

        Args:
            structure: 文件夹结构

        Returns:
            DOCX文件列表
        """
        return get_all_docx_files(structure)

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
        input_base_folder = structure["path"]

        # 逐个处理文件（串行）
        for index, file_info in enumerate(all_files, 1):
            try:
                # 计算输出路径（保持文件夹结构，添加顶层文件夹）
                output_file, relative_path = calculate_output_with_top_folder(
                    file_info["path"],
                    input_base_folder,
                    output_base_dir
                )

                # 转换文件（独立进程）
                success, error = convert_single_file(
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
        批量将DOCX文件转换为PDF（不保持文件夹结构）

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
                from pathlib import Path
                input_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{input_name}.pdf")

                # 转换文件（独立进程）
                success, error = convert_single_file(
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
        """
        获取转换结果摘要

        Args:
            results: 转换结果列表

        Returns:
            摘要字典
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
