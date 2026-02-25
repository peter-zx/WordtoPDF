"""
批量PDF转换页面 - 使用文件夹浏览器对话框
"""

import os
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from services.batch_pdf_service import BatchPDFService


def select_folder_dialog():
    """打开文件夹选择对话框"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes('-topmost', True)  # 置顶

    folder_path = filedialog.askdirectory(
        title="选择文件夹",
        initialdir=os.path.expanduser("~")
    )

    root.destroy()
    return folder_path


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")

    # 初始化session state
    if 'input_folder_path' not in st.session_state:
        st.session_state.input_folder_path = ''
    if 'output_folder_path' not in st.session_state:
        st.session_state.output_folder_path = ''
    if 'folder_structure' not in st.session_state:
        st.session_state.folder_structure = None

    # 步骤1: 选择输入文件夹
    st.markdown("## 📂 步骤1: 选择输入文件夹")

    col1, col2 = st.columns([4, 1])

    with col1:
        input_path = st.text_input(
            "输入文件夹路径",
            value=st.session_state.input_folder_path,
            placeholder="点击右侧按钮选择文件夹...",
            key="input_path_text"
        )

    with col2:
        if st.button("📁 选择文件夹", use_container_width=True, key="browse_input"):
            selected = select_folder_dialog()
            if selected:
                st.session_state.input_folder_path = selected
                st.session_state.folder_structure = None
                st.rerun()

    # 显示当前路径状态
    if st.session_state.input_folder_path:
        if os.path.exists(st.session_state.input_folder_path):
            st.success(f"✅ 已选择: `{st.session_state.input_folder_path}`")
        else:
            st.warning(f"⚠️ 路径不存在: `{st.session_state.input_folder_path}`")

    st.markdown("---")

    # 扫描按钮
    if st.button("🔍 扫描文件夹结构", type="primary", use_container_width=True, disabled=not st.session_state.input_folder_path):
        scan_folder(st.session_state.input_folder_path)

    # 显示扫描结果
    if st.session_state.folder_structure:
        st.markdown("---")
        st.markdown("### 📁 文件夹结构预览")
        display_folder_structure(st.session_state.folder_structure)

    st.markdown("---")

    # 步骤2: 选择输出文件夹
    st.markdown("## 📂 步骤2: 选择输出文件夹")

    col1, col2 = st.columns([4, 1])

    with col1:
        output_path = st.text_input(
            "输出文件夹路径",
            value=st.session_state.output_folder_path,
            placeholder="点击右侧按钮选择文件夹...",
            key="output_path_text"
        )

    with col2:
        if st.button("📁 选择文件夹", use_container_width=True, key="browse_output"):
            selected = select_folder_dialog()
            if selected:
                st.session_state.output_folder_path = selected
                st.rerun()

    # 显示当前路径状态
    if st.session_state.output_folder_path:
        if os.path.exists(st.session_state.output_folder_path):
            st.success(f"✅ 已选择: `{st.session_state.output_folder_path}`")
        else:
            st.warning(f"⚠️ 路径不存在，将自动创建: `{st.session_state.output_folder_path}`")

    st.markdown("---")

    # 步骤3: 开始转换
    st.markdown("## 🚀 步骤3: 开始转换")

    can_convert = (
        st.session_state.folder_structure is not None
        and st.session_state.output_folder_path
    )

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion(
            st.session_state.folder_structure,
            st.session_state.output_folder_path
        )


def scan_folder(folder_path):
    """扫描文件夹"""
    if not folder_path:
        st.error("❌ 请先选择输入文件夹")
        return

    if not os.path.exists(folder_path):
        st.error(f"❌ 文件夹路径不存在: {folder_path}")
        return

    if not os.path.isdir(folder_path):
        st.error(f"❌ 请输入文件夹路径,不是文件路径: {folder_path}")
        return

    with st.spinner("正在扫描文件夹结构..."):
        try:
            structure = BatchPDFService.scan_folder_structure(folder_path)
            st.session_state.folder_structure = structure

            # 统计文件数量
            all_files = BatchPDFService.get_all_docx_files(structure)

            if len(all_files) > 0:
                st.success(f"✅ 扫描完成! 共找到 {len(all_files)} 个DOCX文件")
            else:
                st.warning("⚠️ 未找到DOCX文件,请确认文件夹中包含.docx文件")
        except Exception as e:
            st.error(f"❌ 扫描失败: {str(e)}")


def display_folder_structure(structure, level=0, max_level=3):
    """显示文件夹结构"""
    indent = "  " * level

    # 显示当前文件夹
    folder_name = structure["name"]
    docx_count = len(structure["docx_files"])
    child_count = len(structure["children"])

    if docx_count > 0:
        st.markdown(f"{indent}📁 **{folder_name}** ({docx_count} 个DOCX文件)")
    else:
        st.markdown(f"{indent}📁 {folder_name}")

    # 显示DOCX文件
    for file_info in structure["docx_files"]:
        st.markdown(f"{indent}  📄 `{file_info['name']}`")

    # 递归显示子文件夹
    if level < max_level:
        for child in structure["children"]:
            display_folder_structure(child, level + 1, max_level)
    elif child_count > 0:
        st.markdown(f"{indent}  ... (还有 {child_count} 个子文件夹)")


def execute_conversion(structure, output_dir):
    """执行转换"""
    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        # 执行批量转换
        status_text.text("正在初始化转换...")

        def update_progress(current, total, result):
            progress = current / total
            progress_bar.progress(progress)
            filename = os.path.basename(result['input_file'])
            status_text.text(f"正在转换: {current}/{total} ({progress*100:.1f}%) - {filename}")

        # 执行转换
        results = BatchPDFService.batch_convert_with_structure(
            structure,
            output_dir,
            progress_callback=update_progress
        )

        # 显示结果
        progress_bar.progress(1.0)
        status_text.text("转换完成!")

        display_results(results, output_dir)

    except Exception as e:
        st.error(f"❌ 转换过程中发生错误: {str(e)}")


def display_results(results, output_dir):
    """显示转换结果"""
    # 获取摘要
    summary = BatchPDFService.get_conversion_summary(results)

    # 显示摘要
    st.markdown("---")
    st.markdown("### 📊 转换结果")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("总计", summary["total"])
    with col2:
        st.metric("成功", summary["success"])
    with col3:
        st.metric("失败", summary["failed"])
    with col4:
        st.metric("成功率", summary["success_rate"])

    # 显示详细结果
    if summary["failed"] > 0:
        with st.expander("❌ 查看失败详情", expanded=True):
            for result in results:
                if not result["success"]:
                    error_msg = result.get("error", "未知错误")
                    input_file = os.path.basename(result['input_file'])
                    relative_path = result.get('relative_path', '')
                    if relative_path:
                        st.error(f"`{relative_path}/{input_file}`: {error_msg}")
                    else:
                        st.error(f"`{input_file}`: {error_msg}")

    if summary["success"] > 0:
        with st.expander("✅ 查看成功详情", expanded=False):
            for result in results:
                if result["success"]:
                    input_file = os.path.basename(result['input_file'])
                    output_file = os.path.basename(result['output_file'])
                    relative_path = result.get('relative_path', '')
                    if relative_path:
                        st.success(f"`{relative_path}/{input_file}` → `{output_file}`")
                    else:
                        st.success(f"`{input_file}` → `{output_file}`")

    # 显示输出目录
    st.markdown("---")
    st.info(f"📂 输出目录: `{output_dir}`")

    # 打开输出目录按钮
    col1, col2, col3 = st.columns(3)
    with col2:
        if st.button("📂 打开输出目录", use_container_width=True):
            try:
                os.startfile(output_dir)
            except:
                st.warning("无法自动打开目录，请手动打开")


# 渲染页面
render()
