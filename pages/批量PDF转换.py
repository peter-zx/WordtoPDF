"""
批量PDF转换页面 - 使用独立进程运行tkinter
"""

import os
import streamlit as st
import subprocess
import sys
import json
import tempfile
from services.batch_pdf_service import BatchPDFService
from components.folder_tree import FolderTreeComponent


def select_folder_with_tkinter():
    """使用独立进程运行tkinter选择文件夹"""
    # 创建临时脚本
    script = '''
import tkinter as tk
from tkinter import filedialog
import json
import sys

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

folder_path = filedialog.askdirectory(
    title="选择文件夹",
    initialdir="C:\\\\Users"
)

root.destroy()

# 输出结果
print(json.dumps({"path": folder_path}))
'''
    
    # 写入临时文件
    temp_script = os.path.join(tempfile.gettempdir(), "folder_picker.py")
    with open(temp_script, 'w', encoding='utf-8') as f:
        f.write(script)
    
    # 在独立进程中运行
    try:
        result = subprocess.run(
            [sys.executable, temp_script],
            capture_output=True,
            text=True,
            timeout=60
        )
        if result.returncode == 0 and result.stdout.strip():
            data = json.loads(result.stdout.strip())
            return data.get("path", "")
    except Exception as e:
        pass
    
    return ""


def get_desktop_path():
    """获取桌面路径"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.exists(desktop):
        return desktop
    return os.path.expanduser("~")


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")
    st.info("💡 使用Microsoft Word进行转换，确保已安装Word")

    # 初始化session state
    if 'input_folder_path' not in st.session_state:
        st.session_state.input_folder_path = ''
    if 'output_folder_path' not in st.session_state:
        st.session_state.output_folder_path = os.path.join(get_desktop_path(), "PDF_Output")
    if 'folder_structure' not in st.session_state:
        st.session_state.folder_structure = None

    # 步骤1: 选择输入文件夹
    st.markdown("## 📂 步骤1: 选择输入文件夹")

    col1, col2 = st.columns([4, 1])
    with col1:
        st.text_input(
            "当前文件夹",
            value=st.session_state.input_folder_path,
            key="input_path_display",
            disabled=True
        )
    with col2:
        st.write("")
        if st.button("📁 选择文件夹", use_container_width=True, key="browse_input"):
            folder_path = select_folder_with_tkinter()
            if folder_path:
                st.session_state.input_folder_path = folder_path
                st.session_state.folder_structure = None
                FolderTreeComponent.clear_cache()
                st.rerun()

    if st.session_state.input_folder_path:
        if os.path.exists(st.session_state.input_folder_path):
            st.success(f"✅ `{st.session_state.input_folder_path}`")
        else:
            st.warning("⚠️ 路径不存在")

    # 扫描并显示文件夹结构树
    if st.session_state.input_folder_path and os.path.exists(st.session_state.input_folder_path):
        st.markdown("---")
        st.markdown("### 📁 文件夹结构")

        if st.session_state.folder_structure is None:
            with st.spinner("扫描中..."):
                try:
                    structure = BatchPDFService.scan_folder_structure(
                        st.session_state.input_folder_path
                    )
                    all_files = BatchPDFService.get_all_docx_files(structure)
                    if len(all_files) == 0:
                        st.warning("⚠️ 未找到DOCX文件")
                        return
                    st.session_state.folder_structure = structure
                except Exception as e:
                    st.error(f"❌ 扫描失败: {str(e)}")
                    return

        if st.session_state.folder_structure:
            # 控制按钮
            FolderTreeComponent.render_selection_controls(
                st.session_state.folder_structure
            )

            st.markdown("---")

            # 文件夹树
            FolderTreeComponent.render_folder_tree(st.session_state.folder_structure)

    st.markdown("---")

    # 步骤2: 选择输出文件夹
    st.markdown("## 📂 步骤2: 选择输出文件夹")

    col1, col2 = st.columns([4, 1])
    with col1:
        st.text_input(
            "输出文件夹",
            value=st.session_state.output_folder_path,
            key="output_path_display",
            disabled=True
        )
    with col2:
        st.write("")
        if st.button("📁 选择文件夹", use_container_width=True, key="browse_output"):
            folder_path = select_folder_with_tkinter()
            if folder_path:
                st.session_state.output_folder_path = folder_path
                st.rerun()

    if st.session_state.output_folder_path:
        if os.path.exists(st.session_state.output_folder_path):
            st.success(f"✅ `{st.session_state.output_folder_path}`")
        else:
            st.info(f"📁 将创建")

    st.markdown("---")

    # 步骤3: 开始转换
    st.markdown("## 🚀 步骤3: 开始转换")

    # 获取选中文件
    selected_files = FolderTreeComponent.get_selected_files()
    selected_count = len(selected_files)

    st.write(f"**已选择文件:** {selected_count} 个")

    can_convert = (
        st.session_state.input_folder_path
        and st.session_state.output_folder_path
        and selected_count > 0
    )

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion(
            st.session_state.input_folder_path,
            st.session_state.output_folder_path,
            selected_files
        )


def execute_conversion(input_folder, output_folder, selected_files):
    """执行转换"""
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        file_list = []
        for file_path in selected_files:
            file_list.append({
                "name": os.path.basename(file_path),
                "path": file_path,
                "relative_path": os.path.relpath(file_path, input_folder)
            })

        total = len(file_list)
        status_text.text(f"准备转换 {total} 个文件...")

        def update_progress(current, total, result):
            progress = current / total
            progress_bar.progress(progress)
            filename = os.path.basename(result['input_file'])
            status_text.text(f"转换中: {current}/{total} ({progress*100:.1f}%) - {filename}")

        structure = {
            "name": os.path.basename(input_folder),
            "path": input_folder,
            "type": "folder",
            "children": [],
            "docx_files": file_list
        }

        results = BatchPDFService.batch_convert_with_structure(
            structure,
            output_folder,
            progress_callback=update_progress
        )

        progress_bar.progress(1.0)
        status_text.text("转换完成!")
        display_results(results, output_folder, input_folder)

    except Exception as e:
        st.error(f"❌ 错误: {str(e)}")


def display_results(results, output_dir, input_folder):
    """显示转换结果"""
    summary = BatchPDFService.get_conversion_summary(results)

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

    if summary["failed"] > 0:
        with st.expander("❌ 失败详情", expanded=True):
            for result in results:
                if not result["success"]:
                    error_msg = result.get("error", "未知错误")
                    input_file = os.path.basename(result['input_file'])
                    st.error(f"`{input_file}`: {error_msg}")

    if summary["success"] > 0:
        with st.expander("✅ 成功详情", expanded=False):
            for result in results:
                if result["success"]:
                    input_file = os.path.basename(result['input_file'])
                    output_file = os.path.basename(result['output_file'])
                    st.success(f"`{input_file}` → `{output_file}`")

    st.markdown("---")
    st.info(f"📂 输出目录: `{output_dir}`")

    col1, col2, col3 = st.columns(3)
    with col2:
        if st.button("📂 打开目录", use_container_width=True):
            try:
                os.startfile(output_dir)
            except:
                st.warning("无法打开目录")


render()
