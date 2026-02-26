"""
批量PDF转换页面 - 跨平台兼容版本
- Windows: 支持本地文件夹选择 + 多层文件夹树
- Linux/服务器: 支持文件上传 + 下载
"""

import os
import streamlit as st
import tempfile
import zipfile
import platform
import subprocess
import sys
import json
from datetime import datetime
from services.batch_pdf_service import BatchPDFService
from components.folder_tree import FolderTreeComponent


def is_windows():
    """检查是否为Windows系统"""
    return platform.system() == 'Windows'


def select_folder_with_tkinter():
    """使用独立进程运行tkinter选择文件夹 (Windows专用)"""
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
print(json.dumps({"path": folder_path}))
'''
    
    temp_script = os.path.join(tempfile.gettempdir(), "folder_picker.py")
    with open(temp_script, 'w', encoding='utf-8') as f:
        f.write(script)
    
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
    except:
        pass
    
    return ""


def get_desktop_path():
    """获取桌面路径"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.exists(desktop):
        return desktop
    return os.path.expanduser("~")


def render_windows_mode():
    """Windows模式：支持本地文件夹选择"""
    st.info("💡 使用Microsoft Word/LibreOffice进行转换")
    
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
                        st.warning("⚠️ 未找到Word文件 (DOC/DOCX)")
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

    selected_files = FolderTreeComponent.get_selected_files()
    selected_count = len(selected_files)

    st.write(f"**已选择文件:** {selected_count} 个")

    can_convert = (
        st.session_state.input_folder_path
        and st.session_state.output_folder_path
        and selected_count > 0
    )

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion_windows(
            st.session_state.input_folder_path,
            st.session_state.output_folder_path,
            selected_files
        )


def render_server_mode():
    """服务器模式：支持文件上传"""
    st.info("💡 使用LibreOffice进行转换。请上传Word文件，转换后下载PDF。")
    
    # 初始化session state
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []

    # 步骤1: 上传文件
    st.markdown("## 📂 步骤1: 上传Word文件")
    
    uploaded_files = st.file_uploader(
        "选择Word文件 (DOC/DOCX)",
        type=['doc', 'docx'],
        accept_multiple_files=True,
        help="支持多文件上传"
    )

    if uploaded_files:
        st.success(f"✅ 已选择 {len(uploaded_files)} 个文件")
        for f in uploaded_files:
            st.write(f"  - {f.name} ({f.size / 1024:.1f} KB)")
        st.session_state.uploaded_files = uploaded_files

    st.markdown("---")

    # 步骤2: 开始转换
    st.markdown("## 🚀 步骤2: 开始转换")

    can_convert = len(st.session_state.uploaded_files) > 0

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion_server()

    # 显示结果
    if st.session_state.converted_files:
        st.markdown("---")
        st.markdown("## 📊 转换结果")
        
        success_count = sum(1 for f in st.session_state.converted_files if f['success'])
        total_count = len(st.session_state.converted_files)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("总计", total_count)
        with col2:
            st.metric("成功", success_count)
        with col3:
            st.metric("失败", total_count - success_count)

        # 提供下载
        if success_count > 0:
            st.markdown("---")
            st.markdown("### 📥 下载转换结果")
            
            pdf_files = [f for f in st.session_state.converted_files if f['success']]
            
            if len(pdf_files) == 1:
                pdf_data = pdf_files[0]['data']
                pdf_name = pdf_files[0]['name']
                st.download_button(
                    label=f"📥 下载 {pdf_name}",
                    data=pdf_data,
                    file_name=pdf_name,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            else:
                zip_data = create_zip(pdf_files)
                zip_name = f"converted_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                st.download_button(
                    label=f"📥 下载全部PDF (ZIP, {len(pdf_files)}个文件)",
                    data=zip_data,
                    file_name=zip_name,
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

        # 显示失败详情
        failed_files = [f for f in st.session_state.converted_files if not f['success']]
        if failed_files:
            with st.expander("❌ 查看失败详情", expanded=True):
                for f in failed_files:
                    st.error(f"`{f['name']}`: {f.get('error', '未知错误')}")


def execute_conversion_windows(input_folder, output_folder, selected_files):
    """Windows模式执行转换"""
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


def execute_conversion_server():
    """服务器模式执行转换"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    temp_input = tempfile.mkdtemp(prefix="pdf_input_")
    temp_output = tempfile.mkdtemp(prefix="pdf_output_")
    
    converted_files = []
    
    try:
        total = len(st.session_state.uploaded_files)
        status_text.text(f"准备转换 {total} 个文件...")
        
        for i, uploaded_file in enumerate(st.session_state.uploaded_files):
            input_path = os.path.join(temp_input, uploaded_file.name)
            with open(input_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            
            output_name = os.path.splitext(uploaded_file.name)[0] + '.pdf'
            output_path = os.path.join(temp_output, output_name)
            
            status_text.text(f"转换中: {i+1}/{total} - {uploaded_file.name}")
            
            from services.word_converter import convert_single_file
            success, error = convert_single_file(input_path, output_path)
            
            if success and os.path.exists(output_path):
                with open(output_path, 'rb') as f:
                    pdf_data = f.read()
                converted_files.append({
                    'name': output_name,
                    'success': True,
                    'data': pdf_data
                })
            else:
                converted_files.append({
                    'name': uploaded_file.name,
                    'success': False,
                    'error': error or '转换失败'
                })
            
            progress_bar.progress((i + 1) / total)
        
        st.session_state.converted_files = converted_files
        status_text.text("转换完成!")
        
    except Exception as e:
        st.error(f"❌ 转换过程中发生错误: {str(e)}")
    finally:
        import shutil
        try:
            shutil.rmtree(temp_input, ignore_errors=True)
            shutil.rmtree(temp_output, ignore_errors=True)
        except:
            pass


def create_zip(files):
    """创建ZIP文件"""
    import io
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in files:
            zipf.writestr(f['name'], f['data'])
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


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
        with st.expander("❌ 查看失败详情", expanded=True):
            for result in results:
                if not result["success"]:
                    error_msg = result.get("error", "未知错误")
                    input_file = os.path.basename(result['input_file'])
                    st.error(f"`{input_file}`: {error_msg}")

    if summary["success"] > 0:
        with st.expander("✅ 查看成功详情", expanded=False):
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


def render():
    """渲染批量PDF转换页面 - 根据平台自动选择模式"""
    st.title("📑 批量PDF转换")
    st.markdown("---")
    
    if is_windows():
        render_windows_mode()
    else:
        render_server_mode()


render()
