"""
批量PDF转换页面 - 云端版本
完全重写，修复所有问题
"""

import os
import streamlit as st
import tempfile
import zipfile
import hashlib
import shutil
from datetime import datetime
from services.batch_pdf_service import BatchPDFService
from services.word_converter import convert_single_file


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")
    st.info("💡 使用LibreOffice进行转换。请上传包含Word文件的文件夹(ZIP格式)，转换后下载PDF。")
    
    # 初始化session state
    if 'server_folder_path' not in st.session_state:
        st.session_state.server_folder_path = ''
    if 'server_folder_structure' not in st.session_state:
        st.session_state.server_folder_structure = None
    if 'server_temp_dir' not in st.session_state:
        st.session_state.server_temp_dir = None
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []
    if 'file_selection' not in st.session_state:
        st.session_state.file_selection = {}

    # 步骤1: 上传文件夹(ZIP)
    st.markdown("## 📂 步骤1: 上传文件夹")
    
    st.markdown("### 📦 上传ZIP压缩包")
    st.caption("请将包含Word文件的文件夹压缩成ZIP格式后上传")
    
    uploaded_zip = st.file_uploader(
        "选择ZIP文件",
        type=['zip'],
        help="上传包含DOC/DOCX文件的文件夹压缩包"
    )

    if uploaded_zip:
        st.success(f"✅ 已上传: {uploaded_zip.name} ({uploaded_zip.size / 1024:.1f} KB)")
        
        # 解压并扫描
        if st.session_state.server_temp_dir != uploaded_zip.name:
            with st.spinner("解压并扫描文件夹结构..."):
                try:
                    # 清理旧的临时目录
                    if st.session_state.server_temp_dir and os.path.exists(st.session_state.server_temp_dir):
                        shutil.rmtree(st.session_state.server_temp_dir, ignore_errors=True)
                    
                    # 解压
                    extract_dir = extract_zip_to_temp(uploaded_zip)
                    st.session_state.server_folder_path = extract_dir
                    st.session_state.server_temp_dir = extract_dir
                    
                    # 扫描结构
                    structure = scan_uploaded_folder(extract_dir)
                    all_files = get_all_docx_files(structure)
                    
                    if len(all_files) == 0:
                        st.warning("⚠️ 未找到Word文件 (DOC/DOCX)")
                        st.session_state.server_folder_structure = None
                    else:
                        st.session_state.server_folder_structure = structure
                        # 初始化选择状态
                        init_file_selection(structure)
                        st.success(f"✅ 找到 {len(all_files)} 个Word文件")
                        
                except Exception as e:
                    st.error(f"❌ 解压失败: {str(e)}")
                    st.session_state.server_folder_structure = None

    # 显示文件夹结构树
    if st.session_state.server_folder_structure:
        st.markdown("---")
        st.markdown("### 📁 文件夹结构")
        
        # 控制按钮
        render_selection_controls(st.session_state.server_folder_structure)
        
        st.markdown("---")
        
        # 文件夹树
        render_folder_tree(st.session_state.server_folder_structure)

    st.markdown("---")

    # 步骤2: 开始转换
    st.markdown("## 🚀 步骤2: 开始转换")

    selected_files = get_selected_files()
    selected_count = len(selected_files)

    st.write(f"**已选择文件:** {selected_count} 个")

    can_convert = (
        st.session_state.server_folder_structure
        and selected_count > 0
    )

    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion(selected_files)

    # 显示结果
    if st.session_state.converted_files:
        st.markdown("---")
        st.markdown("## 📊 转换结果")
        
        success_count = sum(1 for f in st.session_state.converted_files if f['success'])
        total_count = len(st.session_state.converted_files)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("总计", total_count)
        with col2:
            st.metric("成功", success_count)
        with col3:
            st.metric("失败", total_count - success_count)
        with col4:
            rate = f"{(success_count/total_count*100):.1f}%" if total_count > 0 else "0%"
            st.metric("成功率", rate)

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
        
        # 显示成功详情
        success_files = [f for f in st.session_state.converted_files if f['success']]
        if success_files:
            with st.expander("✅ 查看成功详情", expanded=False):
                for f in success_files:
                    st.success(f"`{f['original_name']}` → `{f['name']}`")


def extract_zip_to_temp(uploaded_zip):
    """解压上传的ZIP文件到临时目录"""
    temp_dir = tempfile.mkdtemp(prefix="pdf_upload_")
    
    # 保存上传的ZIP文件
    zip_path = os.path.join(temp_dir, "upload.zip")
    with open(zip_path, 'wb') as f:
        f.write(uploaded_zip.getbuffer())
    
    # 解压
    extract_dir = os.path.join(temp_dir, "extracted")
    os.makedirs(extract_dir, exist_ok=True)
    
    with zipfile.ZipFile(zip_path, 'r') as zipf:
        zipf.extractall(extract_dir)
    
    # 删除ZIP文件
    os.remove(zip_path)
    
    return extract_dir


def scan_uploaded_folder(folder_path):
    """扫描上传的文件夹结构"""
    structure = {
        "name": os.path.basename(folder_path),
        "path": folder_path,
        "type": "folder",
        "children": [],
        "docx_files": []
    }
    
    try:
        items = sorted(os.listdir(folder_path))
    except:
        return structure
    
    for item in items:
        item_path = os.path.join(folder_path, item)
        
        if os.path.isdir(item_path):
            child = scan_uploaded_folder(item_path)
            structure["children"].append(child)
        elif item.lower().endswith(('.doc', '.docx')):
            structure["docx_files"].append({
                "name": item,
                "path": item_path,
                "relative_path": item
            })
    
    return structure


def get_all_docx_files(structure):
    """获取所有Word文件"""
    files = []
    for f in structure.get("docx_files", []):
        files.append(f["path"])
    for child in structure.get("children", []):
        files.extend(get_all_docx_files(child))
    return files


def init_file_selection(structure):
    """初始化文件选择状态"""
    all_paths = get_all_docx_files(structure)
    st.session_state.file_selection = {path: False for path in all_paths}


def get_selected_files():
    """获取选中的文件列表"""
    return [path for path, selected in st.session_state.file_selection.items() if selected]


def render_selection_controls(structure):
    """渲染控制按钮"""
    all_paths = get_all_docx_files(structure)
    total = len(all_paths)
    
    selected_count = len(get_selected_files())

    # 注入样式
    st.markdown("""
    <style>
        .selected-count-box {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            padding: 16px 24px;
            text-align: center;
            color: white;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        }
        .selected-count-label {
            font-size: 14px;
            opacity: 0.9;
            margin-bottom: 4px;
        }
        .selected-count-number {
            font-size: 32px;
            font-weight: 700;
            letter-spacing: 2px;
        }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns([1, 1, 1, 1.5])

    with col1:
        if st.button("全选", key="btn_all", use_container_width=True):
            for p in all_paths:
                st.session_state.file_selection[p] = True
            st.rerun()

    with col2:
        if st.button("反选", key="btn_invert", use_container_width=True):
            for p in all_paths:
                st.session_state.file_selection[p] = not st.session_state.file_selection.get(p, False)
            st.rerun()

    with col3:
        if st.button("清空", key="btn_clear", use_container_width=True):
            for p in all_paths:
                st.session_state.file_selection[p] = False
            st.rerun()

    with col4:
        st.markdown(f"""
        <div class="selected-count-box">
            <div class="selected-count-label">已选择文件</div>
            <div class="selected-count-number">{selected_count} / {total}</div>
        </div>
        """, unsafe_allow_html=True)


def render_folder_tree(structure) -> None:
    """渲染文件夹树"""
    # 注入样式
    st.markdown("""
    <style>
        .stCheckbox > label { font-size: 13px !important; }
        .stCheckbox { margin: 0 !important; padding: 0 !important; }
        div[data-testid="stExpander"] { margin: 2px 0 !important; }
        .streamlit-expanderHeader { font-size: 13px !important; padding: 4px !important; }
        .streamlit-expanderContent { padding: 0 0 0 16px !important; border-left: 2px solid #ddd !important; }
    </style>
    """, unsafe_allow_html=True)

    # 渲染根目录文件
    for f in structure.get("docx_files", []):
        render_file(f)

    # 渲染子文件夹
    for idx, child in enumerate(structure.get("children", [])):
        render_folder(child, 0, idx)


def get_stable_key(path: str) -> str:
    """生成稳定的key"""
    return f"cb_{hashlib.md5(path.encode()).hexdigest()[:8]}"


def render_file(file_info) -> None:
    """渲染单个文件"""
    file_path = file_info["path"]
    file_name = file_info["name"]
    
    is_checked = st.session_state.file_selection.get(file_path, False)
    file_key = get_stable_key(file_path)
    
    # 定义callback函数
    def on_file_change():
        st.session_state.file_selection[file_path] = st.session_state[file_key]
    
    # 渲染checkbox
    st.checkbox(
        f"📄 {file_name}", 
        value=is_checked, 
        key=file_key,
        on_change=on_file_change
    )


def render_folder(structure, level: int, idx: int) -> None:
    """渲染文件夹"""
    folder_name = structure["name"]
    folder_path = structure["path"]

    all_paths = get_all_docx_files(structure)
    total = len(all_paths)
    
    selected_count = sum(1 for p in all_paths if st.session_state.file_selection.get(p, False))

    if selected_count == total and total > 0:
        status = "✅"
        folder_checked = True
    elif selected_count > 0:
        status = f"({selected_count}/{total})"
        folder_checked = False
    else:
        status = f"({total})"
        folder_checked = False

    indent = "　" * level
    folder_key = f"folder_{level}_{idx}_{hashlib.md5(folder_path.encode()).hexdigest()[:8]}"
    
    # 定义callback函数
    def on_folder_change():
        new_value = st.session_state[folder_key]
        for p in all_paths:
            st.session_state.file_selection[p] = new_value
    
    # 渲染文件夹checkbox
    st.checkbox(
        f"{indent}📁 {folder_name} {status}",
        value=folder_checked,
        key=folder_key,
        on_change=on_folder_change
    )

    # 渲染内容
    has_content = structure.get("docx_files") or structure.get("children")
    if has_content:
        expanded = selected_count > 0

        with st.expander("▼", expanded=expanded):
            # 渲染文件
            for f in structure.get("docx_files", []):
                render_file(f)

            # 递归渲染子文件夹
            for child_idx, child in enumerate(structure.get("children", [])):
                render_folder(child, level + 1, child_idx)


def execute_conversion(selected_files):
    """执行转换"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    temp_output = tempfile.mkdtemp(prefix="pdf_output_")
    converted_files = []
    
    try:
        total = len(selected_files)
        status_text.text(f"准备转换 {total} 个文件...")
        
        for i, file_path in enumerate(selected_files):
            file_name = os.path.basename(file_path)
            output_name = os.path.splitext(file_name)[0] + '.pdf'
            output_path = os.path.join(temp_output, output_name)
            
            status_text.text(f"转换中: {i+1}/{total} - {file_name}")
            
            success, error = convert_single_file(file_path, output_path)
            
            if success and os.path.exists(output_path):
                with open(output_path, 'rb') as f:
                    pdf_data = f.read()
                converted_files.append({
                    'name': output_name,
                    'original_name': file_name,
                    'success': True,
                    'data': pdf_data
                })
            else:
                converted_files.append({
                    'name': file_name,
                    'original_name': file_name,
                    'success': False,
                    'error': error or '转换失败'
                })
            
            progress_bar.progress((i + 1) / total)
        
        st.session_state.converted_files = converted_files
        status_text.text("转换完成!")
        
    except Exception as e:
        st.error(f"❌ 转换过程中发生错误: {str(e)}")
    finally:
        try:
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


render()
