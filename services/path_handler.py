"""
路径处理工具 - 处理文件夹结构和输出路径
"""

import os
from pathlib import Path


def calculate_output_path(
    input_file_path: str,
    input_base_folder: str,
    output_base_dir: str
) -> str:
    """
    计算输出文件路径，保持文件夹结构

    Args:
        input_file_path: 输入文件完整路径
        input_base_folder: 输入基础文件夹
        output_base_dir: 输出基础目录

    Returns:
        输出PDF文件路径
    """
    # 获取输入文件相对于基础文件夹的路径
    relative_path = os.path.relpath(input_file_path, input_base_folder)

    # 获取文件名（不含扩展名）
    input_name = Path(input_file_path).stem

    # 获取相对路径的目录部分
    relative_dir = os.path.dirname(relative_path)

    # 构建输出目录
    if relative_dir:
        # 有子文件夹，保持结构
        output_dir = os.path.join(output_base_dir, relative_dir)
    else:
        # 在根目录，直接使用输出基础目录
        output_dir = output_base_dir

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 构建输出文件路径
    output_file = os.path.join(output_dir, f"{input_name}.pdf")

    return output_file


def calculate_output_with_top_folder(
    input_file_path: str,
    input_base_folder: str,
    output_base_dir: str
) -> tuple:
    """
    计算输出路径，添加顶层文件夹

    Args:
        input_file_path: 输入文件完整路径
        input_base_folder: 输入基础文件夹
        output_base_dir: 输出基础目录

    Returns:
        (输出文件路径, 相对路径)
    """
    # 获取顶层文件夹名称
    top_folder_name = os.path.basename(input_base_folder)

    # 创建顶层输出目录
    top_output_dir = os.path.join(output_base_dir, top_folder_name)
    os.makedirs(top_output_dir, exist_ok=True)

    # 获取输入文件相对于基础文件夹的路径
    relative_path = os.path.relpath(input_file_path, input_base_folder)

    # 获取文件名（不含扩展名）
    input_name = Path(input_file_path).stem

    # 获取相对路径的目录部分
    relative_dir = os.path.dirname(relative_path)

    # 构建输出目录
    if relative_dir:
        # 有子文件夹，保持结构
        output_dir = os.path.join(top_output_dir, relative_dir)
    else:
        # 在根目录，直接使用顶层目录
        output_dir = top_output_dir

    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)

    # 构建输出文件路径
    output_file = os.path.join(output_dir, f"{input_name}.pdf")

    return output_file, relative_path
