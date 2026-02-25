"""
批量PDF转换页面 - LibreOffice稳定版本
"""

import os
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import subprocess
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


def get_desktop_path():
    """获取桌面路径"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.exists(desktop):
        return desktop
    else:
        return os.path.expanduser("~")


def check_libreoffice():
    """检查LibreOffice是否安装"""
    try:
        # 尝试运行soffice --version
        result = subprocess.run(
            ['soffice', '--version'],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        return result.returncode == 0
    except:
        return False


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")

    # 检查LibreOffice
    libreoffice_installed = check_libreoffice()

    if not libreoffice_installed:
        st.error("❌ 未检测到LibreOffice")
        st.markdown("""
        ### 请先安装LibreOffice

        **为什么需要LibreOffice？**
        - LibreOffice比Word COM接口更稳定
        - 支持批量处理，成功率接近100%
        - 免费开源，无需购买

        **安装方式：**
        1. 访问 https://www.libreoffice.org/download/download/
        2. 下载并安装LibreOffice
        3. 安装完成后刷新页面
        """)
        return

    st.success("✅ LibreOffice已安装，可以开始转换")

    # 初始化session state
    if 'input_folder_path' not in st.session_state:
        st.session_state.input_folder_path = ''
    if 'output_folder_path' not in st.session_state:
        # 默认导出到桌面的PDF_Output文件夹
        desktop = get_desktop_path()
        st.session_state.output_folder_path = os.path.join(desktop, "PDF_Output")
    if 'docx_files' not in st.session_state:
        st.session_state.docx_files = []

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
                st.session_state.docx_files = []
                st.rerun()

    # 显示当前路径状态
    if st.session_state.input_folder_path:
        if os.path.exists(st.session_state.input_folder_path):
            st.success(f"✅ 已选择: `{st.session_state.input_folder_path}`")
        else:
            st.warning(f"⚠️ 路径不存在: `{st.session_state.input_folder_path}`")

    st.markdown("---")

    # 步骤2: 选择输出文件夹
    st.markdown("## 📂 步骤2: 选择输出文件夹")

    col1, col2 = st.columns([4, 1])

    with col1:
        output_path = st.text_input(
            "输出文件夹路径",
            value=st.session_state.output_folder_path,
            placeholder="默认导出到桌面/PDF_Output",
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
            st.info(f"📁 将自动创建: `{st.session_state.output_folder_path}`")

    st.markdown("---")

    # 步骤3: 开始转换
    st.markdown("## 🚀 步骤3: 开始转换")

    can_convert = (
        st.session_state.input_folder_path
        and st.session_state.output_folder_path
    )

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion(
            st.session_state.input_folder_path,
            st.session_state.output_folder_path
        )


def execute_conversion(input_folder, output_folder):
    """执行转换"""
    # 先扫描文件夹
    with st.spinner("正在扫描文件夹..."):
        try:
            structure = BatchPDFService.scan_folder_structure(input_folder)
            all_files = BatchPDFService.get_all_docx_files(structure)

            if len(all_files) == 0:
                st.warning("⚠️ 未找到DOCX文件")
                return

            st.success(f"✅ 找到 {len(all_files)} 个DOCX文件")
        except Exception as e:
            st.error(f"❌ 扫描失败: {str(e)}")
            return

    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        # 执行批量转换
        status_text.text("正在转换文件...")

        def update_progress(current, total, result):
            progress = current / total
            progress_bar.progress(progress)
            filename = os.path.basename(result['input_file'])
            status_text.text(f"正在转换: {current}/{total} ({progress*100:.1f}%) - {filename}")

        # 执行转换（使用LibreOffice，稳定可靠）
        results = BatchPDFService.batch_convert_with_structure(
            structure,
            output_folder,
            progress_callback=update_progress,
            batch_size=10,
            max_workers=2
        )

        # 显示结果
        progress_bar.progress(1.0)
        status_text.text("转换完成!")

        display_results(results, output_folder, input_folder)

    except Exception as e:
        st.error(f"❌ 转换过程中发生错误: {str(e)}")


def display_results(results, output_dir, input_folder):
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

    # 显示输出目录和文件夹结构说明
    st.markdown("---")
    st.info(f"📂 输出目录: `{output_dir}`")
    st.markdown(f"📁 文件夹结构: PDF文件已按原始文件夹结构保存，顶层文件夹名为 `{os.path.basename(input_folder)}`")

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
