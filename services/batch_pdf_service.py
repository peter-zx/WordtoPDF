"""
批量DOCX转PDF服务 - 稳定版本
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
    def convert_with_retry(word_app, input_path: str, output_path: str, max_retries: int = 3) -> Tuple[bool, str]:
        """
        使用Word应用实例转换文件，支持重试

        Args:
            word_app: Word应用实例
            input_path: 输入DOCX文件路径
            output_path: 输出PDF文件路径
            max_retries: 最大重试次数

        Returns:
            (成功标志, 错误信息)
        """
        for attempt in range(max_retries):
            try:
                # 打开DOCX文件
                doc = word_app.Documents.Open(input_path, ReadOnly=True)

                # 保存为PDF格式
                doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF format

                # 关闭文档
                doc.Close()

                return True, ""

            except Exception as e:
                # 如果失败，等待后重试
                if attempt < max_retries - 1:
                    time.sleep(2)  # 等待2秒后重试
                    continue
                else:
                    return False, str(e)

    @staticmethod
    def batch_convert_with_structure(
        structure: Dict,
        output_base_dir: str,
        progress_callback=None,
        batch_size: int = 15
    ) -> List[dict]:
        """
        批量转换DOCX文件,保持文件夹结构（稳定版本）

        Args:
            structure: 文件夹结构
            output_base_dir: 输出基础目录
            progress_callback: 进度回调函数
            batch_size: 每批处理的文件数量

        Returns:
            转换结果列表
        """
        # 获取所有DOCX文件
        all_files = BatchPDFService.get_all_docx_files(structure)
        total = len(all_files)

        if total == 0:
            return []

        results = []
        word_app = None

        # 获取顶层文件夹名称
        top_folder_name = structure["name"]

        try:
            # 创建Word应用实例（只创建一次）
            import win32com.client
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False

            # 创建顶层文件夹
            top_output_dir = os.path.join(output_base_dir, top_folder_name)
            os.makedirs(top_output_dir, exist_ok=True)

            # 分批处理
            for batch_start in range(0, total, batch_size):
                batch_end = min(batch_start + batch_size, total)
                batch_files = all_files[batch_start:batch_end]

                # 处理当前批次
                for i, file_info in enumerate(batch_files):
                    index = batch_start + i + 1

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

                        # 转换文件（带重试）
                        success, error = BatchPDFService.convert_with_retry(
                            word_app,
                            file_info["path"],
                            output_file,
                            max_retries=3
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

                        # 文件间稍作延迟
                        time.sleep(0.3)

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

                # 批次间稍作延迟，让Word休息一下
                if batch_end < total:
                    time.sleep(1)

        finally:
            # 确保关闭Word应用
            if word_app is not None:
                try:
                    word_app.Quit()
                except:
                    pass

        return results

    @staticmethod
    def batch_convert_docx_to_pdf(
        input_files: List[str],
        output_dir: str,
        progress_callback=None,
        batch_size: int = 15
    ) -> List[dict]:
        """
        批量将DOCX文件转换为PDF（稳定版本）

        Args:
            input_files: 输入DOCX文件路径列表
            output_dir: 输出目录
            progress_callback: 进度回调函数
            batch_size: 每批处理的文件数量

        Returns:
            转换结果列表
        """
        results = []
        total = len(input_files)

        if total == 0:
            return []

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        word_app = None

        try:
            # 创建Word应用实例（只创建一次）
            import win32com.client
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False

            # 分批处理
            for batch_start in range(0, total, batch_size):
                batch_end = min(batch_start + batch_size, total)
                batch_files = input_files[batch_start:batch_end]

                # 处理当前批次
                for i, input_file in enumerate(batch_files):
                    index = batch_start + i + 1

                    try:
                        # 生成输出文件名
                        input_name = Path(input_file).stem
                        output_file = os.path.join(output_dir, f"{input_name}.pdf")

                        # 转换文件（带重试）
                        success, error = BatchPDFService.convert_with_retry(
                            word_app,
                            input_file,
                            output_file,
                            max_retries=3
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

                        # 文件间稍作延迟
                        time.sleep(0.3)

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

                # 批次间稍作延迟
                if batch_end < total:
                    time.sleep(1)

        finally:
            # 确保关闭Word应用
            if word_app is not None:
                try:
                    word_app.Quit()
                except:
                    pass

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
