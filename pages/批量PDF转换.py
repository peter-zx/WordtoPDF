"""
批量PDF转换页面
"""

import os
import streamlit as st
import tempfile
from services.batch_pdf_service import BatchPDFService


def get_drives():
    """获取所有驱动器"""
    drives = []
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            drives.append(drive)
    return drives


def get_folder_tree(path, level=0, max_level=2):
    """获取文件夹树形结构"""
    if level > max_level:
        return []
    
    items = []
    try:
        entries = sorted(os.listdir(path))
        for entry in entries:
            full_path = os.path.join(path, entry)
            if os.path.isdir(full_path):
                items.append({
                    "label": f"📁 {entry}",
                    "value": full_path,
                    "children": get_folder_tree(full_path, level + 1, max_level) if level < max_level else []
                })
    except:
        pass
    return items


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")
    
    # 步骤1: 选择文件夹
    st.markdown("## 步骤1: 选择文件夹")
    
    # 选择驱动器
    drives = get_drives()
    selected_drive = st.selectbox(
        "选择驱动器",
        drives,
        index=0,
        key="drive_selector"
    )
    
    # 显示当前路径
    current_path = st.session_state.get('current_path', selected_drive)
    
    # 路径导航
    st.markdown(f"**当前路径:** `{current_path}`")
    
    # 文件夹浏览
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        # 显示当前目录的子文件夹
        try:
            subfolders = []
            entries = sorted(os.listdir(current_path))
            for entry in entries:
                full_path = os.path.join(current_path, entry)
                if os.path.isdir(full_path):
                    subfolders.append((entry, full_path))
            
            if subfolders:
                selected_folder = st.selectbox(
                    "选择文件夹",
                    options=[path for name, path in subfolders],
                    format_func=lambda x: [name for name, path in subfolders if path == x][0],
                    key="folder_selector"
                )
            else:
                st.info("当前目录没有子文件夹")
                selected_folder = None
        except Exception as e:
            st.error(f"无法访问目录: {str(e)}")
            selected_folder = None
    
    with col2:
        if st.button("📂 进入", use_container_width=True):
            if selected_folder:
                st.session_state.current_path = selected_folder
                st.rerun()
    
    with col3:
        if st.button("⬆️ 上级", use_container_width=True):
            parent = os.path.dirname(current_path)
            if parent and parent != current_path:
                st.session_state.current_path = parent
                st.rerun()
    
    # 扫描按钮
    st.markdown("---")
    if st.button("🔍 扫描当前文件夹", type="primary", use_container_width=True):
        scan_folder(current_path)
    
    # 显示扫描结果
    if "folder_structure" in st.session_state:
        st.markdown("---")
        st.markdown("### 文件夹结构")
        display_folder_structure(st.session_state.folder_structure)
    
    st.markdown("---")
    
    # 步骤2: 选择输出目录
    st.markdown("## 步骤2: 选择输出目录")
    
    # 输出目录选择
    output_drive = st.selectbox(
        "选择输出驱动器",
        drives,
        index=0,
        key="output_drive_selector"
    )
    
    output_path = st.session_state.get('output_path', output_drive)
    
    # 输出路径导航
    st.markdown(f"**输出路径:** `{output_path}`")
    
    # 输出文件夹浏览
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        try:
            output_subfolders = []
            entries = sorted(os.listdir(output_path))
            for entry in entries:
                full_path = os.path.join(output_path, entry)
                if os.path.isdir(full_path):
                    output_subfolders.append((entry, full_path))
            
            if output_subfolders:
                selected_output_folder = st.selectbox(
                    "选择输出文件夹",
                    options=[path for name, path in output_subfolders],
                    format_func=lambda x: [name for name, path in output_subfolders if path == x][0],
                    key="output_folder_selector"
                )
            else:
                st.info("当前目录没有子文件夹")
                selected_output_folder = None
        except Exception as e:
            st.error(f"无法访问目录: {str(e)}")
            selected_output_folder = None
    
    with col2:
        if st.button("📂 进入", key="output_enter", use_container_width=True):
            if selected_output_folder:
                st.session_state.output_path = selected_output_folder
                st.rerun()
    
    with col3:
        if st.button("⬆️ 上级", key="output_up", use_container_width=True):
            parent = os.path.dirname(output_path)
            if parent and parent != output_path:
                st.session_state.output_path = parent
                st.rerun()
    
    # 使用当前路径作为输出目录
    output_dir = output_path
    
    st.markdown("---")
    
    # 步骤3: 开始转换
    st.markdown("## 步骤3: 开始转换")
    
    can_convert = "folder_structure" in st.session_state and output_dir
    
    if st.button("🚀 开始转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_folder_conversion(st.session_state.folder_structure, output_dir)


def scan_folder(folder_path):
    """扫描文件夹"""
    if not os.path.exists(folder_path):
        st.error(f"❌ 文件夹路径不存在: {folder_path}")
        return
    
    if not os.path.isdir(folder_path):
        st.error(f"❌ 请输入文件夹路径,不是文件路径: {folder_path}")
        return
    
    with st.spinner("正在扫描文件夹..."):
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


def display_folder_structure(structure, level=0):
    """显示文件夹结构"""
    indent = "  " * level
    
    # 显示当前文件夹
    folder_name = structure["name"]
    docx_count = len(structure["docx_files"])
    
    if docx_count > 0:
        st.markdown(f"{indent}📁 **{folder_name}** ({docx_count} 个DOCX文件)")
    else:
        st.markdown(f"{indent}📁 {folder_name}")
    
    # 显示DOCX文件
    for file_info in structure["docx_files"]:
        st.markdown(f"{indent}  📄 {file_info['name']}")
    
    # 递归显示子文件夹
    for child in structure["children"]:
        display_folder_structure(child, level + 1)


def execute_folder_conversion(structure, output_dir):
    """执行文件夹转换"""
    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # 执行批量转换
        status_text.text("正在转换文件...")
        
        def update_progress(current, total, result):
            progress = current / total
            progress_bar.progress(progress)
            status_text.text(f"正在转换: {current}/{total} ({progress*100:.1f}%) - {os.path.basename(result['input_file'])}")
        
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
    st.markdown("### 转换结果")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("总计", summary["total"])
    with col2:
        st.metric("成功", summary["success"])
    with col3:
        st.metric("失败", summary["failed"])
    
    # 显示详细结果
    if summary["failed"] > 0:
        with st.expander("❌ 查看失败详情", expanded=True):
            for result in results:
                if not result["success"]:
                    st.error(f"{os.path.basename(result['input_file'])}: {result['error']}")
    
    if summary["success"] > 0:
        with st.expander("✅ 查看成功详情", expanded=False):
            for result in results:
                if result["success"]:
                    st.success(f"{os.path.basename(result['input_file'])} → {os.path.basename(result['output_file'])}")
    
    # 显示输出目录
    st.info(f"📂 输出目录: {output_dir}")


# 渲染页面
render()
