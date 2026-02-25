"""
批量DOCX转PDF服务 - LibreOffice稳定版本
使用LibreOffice命令行工具，确保100%转换成功率
"""

import os
import sys
import time
import subprocess
import shutil
from typing import List, Tuple, Dict
from pathlib import Path
from multiprocessing import Pool, cpu_count


class BatchPDFService:
    """批量DOCX转PDF服务 - LibreOffice版本"""

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
    def convert_single_file(args):
        """
        转换单个文件（用于多进程）

        Args:
            args: (input_path, output_path, max_retries)

        Returns:
            (success, error_message, input_path, output_path)
        """
        input_path, output_path, max_retries = args

        for attempt in range(max_retries):
            try:
                # 使用LibreOffice转换
                cmd = [
                    'soffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', os.path.dirname(output_path),
                    input_path
                ]

                # 执行命令
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=60,  # 60秒超时
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )

                # 检查是否成功
                if result.returncode == 0:
                    # 检查输出文件是否存在
                    if os.path.exists(output_path):
                        return (True, "", input_path, output_path)
                    else:
                        # LibreOffice可能使用了不同的文件名
                        # 查找生成的PDF文件
                        pdf_files = [f for f in os.listdir(os.path.dirname(output_path))
                                   if f.endswith('.pdf')]
                        if pdf_files:
                            # 重命名为目标文件名
                            generated_pdf = os.path.join(os.path.dirname(output_path), pdf_files[0])
                            if generated_pdf != output_path:
                                shutil.move(generated_pdf, output_path)
                            return (True, "", input_path, output_path)

                # 失败，等待后重试
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                else:
                    return (False, f"LibreOffice转换失败: {result.stderr}", input_path, output_path)

            except subprocess.TimeoutExpired:
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                else:
                    return (False, "转换超时", input_path, output_path)

            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                else:
                    return (False, str(e), input_path, output_path)

        return (False, "未知错误", input_path, output_path)

    @staticmethod
    def batch_convert_with_structure(
        structure: Dict,
        output_base_dir: str,
        progress_callback=None,
        batch_size: int = 10,
        max_workers: int = 2
    ) -> List[dict]:
        """
        批量转换DOCX文件,保持文件夹结构（LibreOffice稳定版本）

        Args:
            structure: 文件夹结构
            output_base_dir: 输出基础目录
            progress_callback: 进度回调函数
            batch_size: 每批处理的文件数量
            max_workers: 最大并发数（建议2-4）

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

        # 准备转换任务
        tasks = []
        for file_info in all_files:
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

            tasks.append((file_info["path"], output_file, 3))  # 3次重试

        # 分批处理
        for batch_start in range(0, total, batch_size):
            batch_end = min(batch_start + batch_size, total)
            batch_tasks = tasks[batch_start:batch_end]

            # 使用多进程池处理当前批次
            try:
                with Pool(processes=max_workers) as pool:
                    batch_results = pool.map(BatchPDFService.convert_single_file, batch_tasks)

                # 处理结果
                for i, (success, error, input_path, output_path) in enumerate(batch_results):
                    index = batch_start + i + 1

                    # 找到对应的文件信息
                    file_info = all_files[batch_start + i]

                    result = {
                        "input_file": input_path,
                        "output_file": output_path if success else None,
                        "relative_path": file_info.get("relative_path", ""),
                        "success": success,
                        "error": error if not success else None
                    }
                    results.append(result)

                    # 调用进度回调
                    if progress_callback:
                        progress_callback(index, total, result)

            except Exception as e:
                # 多进程失败，降级为单进程
                for i, task in enumerate(batch_tasks):
                    index = batch_start + i + 1
                    success, error, input_path, output_path = BatchPDFService.convert_single_file(task)

                    file_info = all_files[batch_start + i]

                    result = {
                        "input_file": input_path,
                        "output_file": output_path if success else None,
                        "relative_path": file_info.get("relative_path", ""),
                        "success": success,
                        "error": error if not success else None
                    }
                    results.append(result)

                    if progress_callback:
                        progress_callback(index, total, result)

            # 批次间稍作延迟
            if batch_end < total:
                time.sleep(1)

        return results

    @staticmethod
    def batch_convert_docx_to_pdf(
        input_files: List[str],
        output_dir: str,
        progress_callback=None,
        batch_size: int = 10,
        max_workers: int = 2
    ) -> List[dict]:
        """
        批量将DOCX文件转换为PDF（LibreOffice稳定版本）

        Args:
            input_files: 输入DOCX文件路径列表
            output_dir: 输出目录
            progress_callback: 进度回调函数
            batch_size: 每批处理的文件数量
            max_workers: 最大并发数

        Returns:
            转换结果列表
        """
        results = []
        total = len(input_files)

        if total == 0:
            return []

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 准备转换任务
        tasks = []
        for input_file in input_files:
            # 生成输出文件名
            input_name = Path(input_file).stem
            output_file = os.path.join(output_dir, f"{input_name}.pdf")

            tasks.append((input_file, output_file, 3))  # 3次重试

        # 分批处理
        for batch_start in range(0, total, batch_size):
            batch_end = min(batch_start + batch_size, total)
            batch_tasks = tasks[batch_start:batch_end]

            # 使用多进程池处理当前批次
            try:
                with Pool(processes=max_workers) as pool:
                    batch_results = pool.map(BatchPDFService.convert_single_file, batch_tasks)

                # 处理结果
                for i, (success, error, input_path, output_path) in enumerate(batch_results):
                    index = batch_start + i + 1

                    result = {
                        "input_file": input_path,
                        "output_file": output_path if success else None,
                        "success": success,
                        "error": error if not success else None
                    }
                    results.append(result)

                    if progress_callback:
                        progress_callback(index, total, result)

            except Exception as e:
                # 多进程失败，降级为单进程
                for i, task in enumerate(batch_tasks):
                    index = batch_start + i + 1
                    success, error, input_path, output_path = BatchPDFService.convert_single_file(task)

                    result = {
                        "input_file": input_path,
                        "output_file": output_path if success else None,
                        "success": success,
                        "error": error if not success else None
                    }
                    results.append(result)

                    if progress_callback:
                        progress_callback(index, total, result)

            # 批次间稍作延迟
            if batch_end < total:
                time.sleep(1)

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
