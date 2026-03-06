# -*- coding: utf-8 -*-
"""文件夹管理模块"""
import os


class FolderManager:
    """文件夹管理器"""

    def __init__(self, base_path=""):
        self.base_path = base_path
        self.structure = {}

    def set_base_path(self, path):
        """设置基础路径"""
        self.base_path = path

    def set_structure(self, structure):
        """设置文件夹结构"""
        self.structure = structure

    def create_structure(self, root_name=""):
        """创建文件夹结构"""
        if not self.structure:
            raise ValueError("文件夹结构为空")

        if root_name:
            target_root = os.path.join(self.base_path, root_name)
        else:
            target_root = self.base_path

        folder_count = 0

        for level1, level2_dict in self.structure.items():
            path1 = os.path.join(target_root, level1)
            os.makedirs(path1, exist_ok=True)
            folder_count += 1

            if isinstance(level2_dict, dict):
                for level2, level3_list in level2_dict.items():
                    path2 = os.path.join(path1, level2)
                    os.makedirs(path2, exist_ok=True)
                    folder_count += 1

                    if level3_list:
                        for level3 in level3_list:
                            path3 = os.path.join(path2, level3)
                            os.makedirs(path3, exist_ok=True)
                            folder_count += 1

        return target_root, folder_count

    def get_all_subfolders(self, folder_path):
        """获取文件夹下所有子文件夹"""
        folders = [folder_path]

        for root, dirs, files in os.walk(folder_path):
            for d in dirs:
                folders.append(os.path.join(root, d))

        return folders

    def find_matching_folder(self, file_name, folder_list):
        """根据文件名查找匹配的文件夹"""
        file_name_lower = file_name.lower()

        for folder in folder_list:
            folder_name = os.path.basename(folder).lower()
            if folder_name in file_name_lower:
                return folder

        return None

    def folder_exists(self, path):
        """检查文件夹是否存在"""
        return os.path.exists(path) and os.path.isdir(path)

    def count_folders(self, folder_path):
        """统计文件夹数量"""
        count = 0
        for root, dirs, files in os.walk(folder_path):
            count += len(dirs)
        return count
