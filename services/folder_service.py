"""
文件夹生成服务模块
"""

import os
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any


class FolderService:
    """文件夹生成服务"""

    @staticmethod
    def parse_text_to_names(text: str) -> List[str]:
        """
        解析文本为文件夹名称列表
        支持多级文件夹结构识别

        Args:
            text: 输入文本

        Returns:
            文件夹名称列表(支持多级,如: 工具/钓鱼竿工具)
        """
        if not text:
            return []

        lines = text.strip().split('\n')
        names = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 尝试识别分隔符
            # 支持的格式: "工具 - 钓鱼竿工具" 或 "工具/钓鱼竿工具" 或 "工具|钓鱼竿工具"
            separators = [' - ', ' -', '-', ' / ', '/', ' | ', '|', ' > ', '>', ' → ', '→']

            # 找到第一个出现的分隔符
            earliest_pos = -1
            used_separator = None

            for sep in separators:
                pos = line.find(sep)
                if pos != -1 and (earliest_pos == -1 or pos < earliest_pos):
                    earliest_pos = pos
                    used_separator = sep

            if used_separator:
                # 使用分隔符分割,然后重新组合为多级路径
                parts = line.split(used_separator)
                parts = [p.strip() for p in parts if p.strip()]

                # 如果有多个部分,生成多级路径
                if len(parts) >= 2:
                    # 第一部分是父文件夹
                    parent = parts[0]
                    # 后续部分都是子文件夹
                    for child in parts[1:]:
                        full_path = f"{parent}/{child}"
                        names.append(full_path)
                else:
                    names.append(line)
            else:
                # 没有分隔符,直接作为文件夹名称
                names.append(line)

        return names

    @staticmethod
    def read_excel_file(file_path: str) -> List[str]:
        """
        读取Excel文件获取文件夹名称
        支持多列数据,每行创建一个文件夹路径

        Args:
            file_path: Excel文件路径

        Returns:
            文件夹名称列表
        """
        try:
            import pandas as pd

            # 读取Excel
            df = pd.read_excel(file_path)

            # 获取所有列
            folder_paths = []

            # 遍历每一行
            for index, row in df.iterrows():
                # 获取非空值
                parts = []
                for col in df.columns:
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        parts.append(str(value).strip())

                # 如果有数据,组合成路径
                if parts:
                    folder_path = '/'.join(parts)
                    folder_paths.append(folder_path)

            return folder_paths

        except ImportError:
            raise Exception("需要安装pandas和openpyxl: pip install pandas openpyxl")
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {str(e)}")

    @staticmethod
    def read_csv_file(file_path: str) -> List[str]:
        """
        读取CSV文件获取文件夹名称
        支持多列数据,每行创建一个文件夹路径

        Args:
            file_path: CSV文件路径

        Returns:
            文件夹名称列表
        """
        try:
            import pandas as pd

            # 读取CSV
            df = pd.read_csv(file_path)

            # 获取所有列
            folder_paths = []

            # 遍历每一行
            for index, row in df.iterrows():
                # 获取非空值
                parts = []
                for col in df.columns:
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        parts.append(str(value).strip())

                # 如果有数据,组合成路径
                if parts:
                    folder_path = '/'.join(parts)
                    folder_paths.append(folder_path)

            return folder_paths

        except ImportError:
            raise Exception("需要安装pandas: pip install pandas")
        except Exception as e:
            raise Exception(f"读取CSV文件失败: {str(e)}")

    @staticmethod
    def read_text_file(file_path: str) -> List[str]:
        """
        读取文本文件获取文件夹名称

        Args:
            file_path: 文本文件路径

        Returns:
            文件夹名称列表
        """
        encodings = ['utf-8', 'gbk', 'utf-8-sig', 'latin-1']

        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    text = f.read()
                return FolderService.parse_text_to_names(text)
            except (UnicodeDecodeError, UnicodeError):
                continue

        raise Exception("无法识别文件编码")

    @staticmethod
    def create_folders_with_files(
        folder_names: List[str],
        target_dir: str,
        template_files: List[str],
        use_timestamp: bool = True
    ) -> Dict[str, Any]:
        """
        创建文件夹并复制文件

        Args:
            folder_names: 文件夹名称列表
            target_dir: 目标目录
            template_files: 模板文件路径列表
            use_timestamp: 是否使用时间戳创建顶层文件夹

        Returns:
            操作结果
        """
        results = {
            'success': [],
            'failed': [],
            'skipped': [],
            'total': len(folder_names)
        }

        try:
            # 创建顶层目录
            if use_timestamp:
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                base_dir = os.path.join(target_dir, timestamp)
            else:
                base_dir = target_dir

            os.makedirs(base_dir, exist_ok=True)

            # 创建每个文件夹
            for folder_name in folder_names:
                try:
                    folder_path = os.path.join(base_dir, folder_name)

                    # 检查是否已存在
                    if os.path.exists(folder_path):
                        results['skipped'].append(folder_name)
                        continue

                    # 创建文件夹
                    os.makedirs(folder_path, exist_ok=True)

                    # 复制模板文件
                    copied_files = []
                    for template_file in template_files:
                        if os.path.exists(template_file):
                            file_name = os.path.basename(template_file)
                            dest_path = os.path.join(folder_path, file_name)

                            # 处理重复文件名
                            counter = 1
                            while os.path.exists(dest_path):
                                name, ext = os.path.splitext(file_name)
                                dest_path = os.path.join(folder_path, f"{name}_{counter}{ext}")
                                counter += 1

                            shutil.copy2(template_file, dest_path)
                            copied_files.append(os.path.basename(dest_path))

                    results['success'].append({
                        'folder': folder_name,
                        'path': folder_path,
                        'files': copied_files
                    })

                except Exception as e:
                    results['failed'].append({
                        'folder': folder_name,
                        'error': str(e)
                    })

        except Exception as e:
            return {
                'error': f'创建顶层目录失败: {str(e)}'
            }

        return results

    @staticmethod
    def scan_folder_files(folder_path: str, recursive: bool = False) -> List[Dict[str, Any]]:
        """
        扫描文件夹中的文件

        Args:
            folder_path: 文件夹路径
            recursive: 是否递归扫描

        Returns:
            文件列表
            [
                {
                    'name': '文件名',
                    'path': '完整路径',
                    'size': 大小,
                    'ext': '扩展名'
                }
            ]
        """
        files = []

        if not os.path.exists(folder_path):
            return files

        if recursive:
            for root, dirs, filenames in os.walk(folder_path):
                for filename in filenames:
                    file_path = os.path.join(root, filename)
                    files.append({
                        'name': filename,
                        'path': file_path,
                        'size': os.path.getsize(file_path),
                        'ext': Path(filename).suffix.lower()
                    })
        else:
            for item in os.listdir(folder_path):
                item_path = os.path.join(folder_path, item)
                if os.path.isfile(item_path):
                    files.append({
                        'name': item,
                        'path': item_path,
                        'size': os.path.getsize(item_path),
                        'ext': Path(item).suffix.lower()
                    })

        return files
