"""
文件处理服务模块
"""

import os
import shutil
from typing import List, Dict, Any


class FileService:
    """文件处理服务"""

    @staticmethod
    def save_uploaded_files(uploaded_files, temp_dir: str) -> List[str]:
        """
        保存上传的文件到临时目录

        Args:
            uploaded_files: 上传的文件列表
            temp_dir: 临时目录

        Returns:
            保存的文件路径列表
        """
        os.makedirs(temp_dir, exist_ok=True)
        file_paths = []

        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            file_paths.append(file_path)

        return file_paths

    @staticmethod
    def cleanup_temp_directory(temp_dir: str) -> bool:
        """
        清理临时目录

        Args:
            temp_dir: 临时目录

        Returns:
            是否清理成功
        """
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return True
        except Exception:
            return False

    @staticmethod
    def copy_file_to_folders(
        source_file: str,
        target_folders: List[str],
        handle_duplicates: bool = True
    ) -> Dict[str, Any]:
        """
        将文件复制到多个文件夹

        Args:
            source_file: 源文件路径
            target_folders: 目标文件夹列表
            handle_duplicates: 是否处理重复文件名

        Returns:
            复制结果
        """
        results = {
            'success': [],
            'failed': []
        }

        if not os.path.exists(source_file):
            return {'error': '源文件不存在'}

        file_name = os.path.basename(source_file)

        for folder_path in target_folders:
            try:
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path, exist_ok=True)

                dest_path = os.path.join(folder_path, file_name)

                # 处理重复文件名
                if handle_duplicates and os.path.exists(dest_path):
                    name, ext = os.path.splitext(file_name)
                    counter = 1
                    while os.path.exists(dest_path):
                        dest_path = os.path.join(folder_path, f"{name}_{counter}{ext}")
                        counter += 1

                shutil.copy2(source_file, dest_path)

                results['success'].append({
                    'folder': folder_path,
                    'dest_file': dest_path
                })

            except Exception as e:
                results['failed'].append({
                    'folder': folder_path,
                    'error': str(e)
                })

        return results

    @staticmethod
    def get_file_info(file_path: str) -> Dict[str, Any]:
        """
        获取文件信息

        Args:
            file_path: 文件路径

        Returns:
            文件信息
        """
        if not os.path.exists(file_path):
            return {'error': '文件不存在'}

        return {
            'name': os.path.basename(file_path),
            'path': file_path,
            'size': os.path.getsize(file_path),
            'ext': os.path.splitext(file_path)[1].lower(),
            'created': os.path.getctime(file_path),
            'modified': os.path.getmtime(file_path)
        }

    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """
        格式化文件大小

        Args:
            size_bytes: 字节数

        Returns:
            格式化的大小字符串
        """
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"
