"""
文件夹扫描工具 - 专门处理文件夹结构
"""

import os
from typing import List, Dict


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
                child_structure = scan_folder_structure(item_path)
                structure["children"].append(child_structure)

            elif item.lower().endswith('.docx'):
                # 添加DOCX文件，记录相对路径
                structure["docx_files"].append({
                    "name": item,
                    "path": item_path,
                    "relative_path": os.path.relpath(item_path, folder_path)
                })

    except Exception as e:
        print(f"扫描文件夹失败: {str(e)}")

    return structure


def get_all_docx_files(structure: Dict) -> List[Dict]:
    """
    从文件夹结构中获取所有DOCX文件

    Args:
        structure: 文件夹结构字典

    Returns:
        DOCX文件列表
    """
    files = []

    # 添加当前文件夹的DOCX文件
    files.extend(structure.get("docx_files", []))

    # 递归添加子文件夹的DOCX文件
    for child in structure.get("children", []):
        files.extend(get_all_docx_files(child))

    return files


def print_folder_structure(structure: Dict, level: int = 0):
    """
    打印文件夹结构（用于调试）

    Args:
        structure: 文件夹结构字典
        level: 缩进级别
    """
    indent = "  " * level
    print(f"{indent}📁 {structure['name']}/")

    # 打印当前文件夹的DOCX文件
    for docx in structure.get("docx_files", []):
        print(f"{indent}  📄 {docx['name']} - {docx['relative_path']}")

    # 递归打印子文件夹
    for child in structure.get("children", []):
        print_folder_structure(child, level + 1)
