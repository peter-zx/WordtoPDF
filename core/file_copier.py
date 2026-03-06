# -*- coding: utf-8 -*-
"""文件复制模块"""
import os
import shutil
from .logger import logger


class FileCopier:
    """文件复制器"""

    def __init__(self):
        self.scanned_files = []
        self.check_states = {}

    def scan_folder(self, folder_path, file_types=None):
        """扫描文件夹获取文件列表"""
        logger.info(f"开始扫描文件夹: {folder_path}")
        
        if not os.path.exists(folder_path):
            error_msg = f"文件夹不存在: {folder_path}"
            logger.error(error_msg)
            raise FileNotFoundError(error_msg)

        if file_types is None:
            file_types = [".docx", ".doc", ".pdf"]

        self.scanned_files = []
        self.check_states = {}
        file_count = 0

        for root, dirs, files in os.walk(folder_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in file_types:
                    full_path = os.path.join(root, file)
                    size = os.path.getsize(full_path)

                    idx = len(self.scanned_files)
                    self.scanned_files.append({
                        "path": full_path,
                        "name": file,
                        "size": size,
                        "size_str": self._format_size(size),
                        "type": ext,
                        "selected": False
                    })
                    self.check_states[idx] = False
                    file_count += 1

        logger.info(f"扫描完成，找到 {file_count} 个文件")
        return self.scanned_files

    def get_selected_files(self):
        """获取选中的文件"""
        return [f for f in self.scanned_files if f.get("selected", False)]

    def select_file(self, index, selected=True):
        """设置文件选中状态"""
        if 0 <= index < len(self.scanned_files):
            self.check_states[index] = selected
            self.scanned_files[index]["selected"] = selected

    def select_all(self):
        """全选"""
        for i in range(len(self.scanned_files)):
            self.check_states[i] = True
            self.scanned_files[i]["selected"] = True

    def deselect_all(self):
        """取消全选"""
        for i in range(len(self.scanned_files)):
            self.check_states[i] = False
            self.scanned_files[i]["selected"] = False

    def get_selected_count(self):
        """获取选中数量"""
        return sum(1 for v in self.check_states.values() if v)

    def copy_to_folder(self, files, target_folder, folder_list=None):
        """复制文件到目标文件夹"""
        results = []
        success = 0
        fail = 0

        for file_info in files:
            file_path = file_info["path"]
            file_name = file_info["name"]

            target_path = None

            # 如果有文件夹列表，尝试匹配
            if folder_list:
                for folder in folder_list:
                    folder_name = os.path.basename(folder).lower()
                    if folder_name in file_name.lower():
                        target_path = os.path.join(folder, file_name)
                        break

            # 如果没有匹配到，使用目标文件夹
            if not target_path:
                target_path = os.path.join(target_folder, file_name)

            try:
                shutil.copy2(file_path, target_path)
                success += 1
                rel_path = os.path.relpath(os.path.dirname(target_path), target_folder)
                results.append(f"✓ {file_name} -> {rel_path if rel_path != '.' else '根目录'}")
            except Exception as e:
                fail += 1
                results.append(f"✗ {file_name} 失败: {str(e)}")

        return results, success, fail

    def copy_single(self, file_path, target_folder):
        """复制单个文件"""
        file_name = os.path.basename(file_path)
        target_path = os.path.join(target_folder, file_name)
        shutil.copy2(file_path, target_path)
        return target_path

    def copy_to_selected_leaf_folders(self, files, selected_folder):
        """将文件复制到选中文件夹下的所有最底层文件夹（叶子节点）"""
        logger.info(f"开始复制文件到: {selected_folder}")
        logger.info(f"文件数量: {len(files)}")
        
        # 获取选中文件夹下的所有最底层文件夹（没有子文件夹的文件夹）
        leaf_folders = self._get_all_leaf_folders(selected_folder)
        logger.info(f"找到 {len(leaf_folders)} 个最底层文件夹")
        
        results = []
        success = 0
        fail = 0
        
        for file_info in files:
            file_name = file_info["name"]
            logger.debug(f"处理文件: {file_name}")
            
            # 复制到所有最底层文件夹（不包括选中文件夹本身）
            for leaf_folder in leaf_folders:
                try:
                    target_path = os.path.join(leaf_folder, file_name)
                    shutil.copy2(file_info["path"], target_path)
                    success += 1
                    rel_path = os.path.relpath(leaf_folder, selected_folder)
                    results.append(f"✓ {file_name} -> {rel_path}")
                    logger.debug(f"复制成功: {file_name} -> {rel_path}")
                except Exception as e:
                    fail += 1
                    error_msg = f"✗ {file_name} -> {leaf_folder}: {str(e)}"
                    results.append(error_msg)
                    logger.error(f"复制失败: {error_msg}")
        
        logger.info(f"复制完成: 成功 {success}, 失败 {fail}")
        return results, success, fail

    def _get_all_leaf_folders(self, folder_path):
        """获取文件夹下的所有最底层文件夹（没有子文件夹的文件夹）"""
        leaf_folders = []
        
        def find_leaves(current_path):
            has_subfolders = False
            try:
                for item in os.listdir(current_path):
                    item_path = os.path.join(current_path, item)
                    if os.path.isdir(item_path):
                        has_subfolders = True
                        find_leaves(item_path)
                
                # 如果没有子文件夹，就是最底层文件夹
                if not has_subfolders:
                    leaf_folders.append(current_path)
            except PermissionError:
                pass
        
        find_leaves(folder_path)
        return leaf_folders

    def _get_direct_subfolders(self, folder_path):
        """获取文件夹下的直接子文件夹（不包括子文件夹的子文件夹）"""
        subfolders = []
        try:
            for item in os.listdir(folder_path):
                item_path = os.path.join(folder_path, item)
                if os.path.isdir(item_path):
                    subfolders.append(item_path)
        except PermissionError:
            pass
        return subfolders

    @staticmethod
    def _format_size(size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"

    def clear(self):
        """清空文件列表"""
        self.scanned_files = []
        self.check_states = {}
