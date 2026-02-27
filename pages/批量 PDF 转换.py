"""
批量 PDF 转换页面 - 实用版本
功能:
1. 上传 ZIP 压缩包，保持文件夹结构
2. 树形显示文件，支持展开/折叠
3. 多选文件 (全选/反选/清空)
4. 转换后保持原始文件夹结构打包下载
"""

import os
import streamlit as st
import tempfile
import zipfile
import shutil
from datetime import datetime
from services.word_converter import convert_single_file


def render():
    """渲染批量 PDF 转换页面"""
    st.title("📑 批量 PDF 转换")
    st.markdown("---")
    
    # 初始化 session state
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'file_selection' not in st.session_state:
        st.session_state.file_selection = {}
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []
    if 'folder_tree' not in st.session_state:
        st.session_state.folder_tree = {}
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = None
    if 'expand_all' not in st.session_state:
        st.session_state.expand_all = False
    
    # ========== 步骤 1: 上传 ZIP 压缩包 ==========
    st.markdown("### 📂 步骤 1: 上传文件夹 (ZIP 压缩包)")
    
    st.info("""
💡 **使用说明**:
1. 将包含 Word 文件的整个文件夹压缩为 **ZIP 格式**
2. 系统会保持原始文件夹结构显示所有文件
3. 支持多层嵌套文件夹
""")
    
    uploaded_zip = st.file_uploader(
        "点击选择 ZIP 文件上传",
        type=['zip'],
        help="上传包含 DOC/DOCX 文件的文件夹压缩包"
    )
    
    if uploaded_zip:
        st.success(f"✅ 已上传：{uploaded_zip.name} ({uploaded_zip.size / 1024:.1f} KB)")
        
        with st.spinner("📦 正在解压并分析文件夹结构..."):
            try:
                if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
                    shutil.rmtree(st.session_state.temp_dir, ignore_errors=True)
                
                temp_dir = tempfile.mkdtemp(prefix="pdf_input_")
                st.session_state.temp_dir = temp_dir
                
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                word_exts = ['.doc', '.docx']
                all_files = []
                folder_tree = {'_root': {'files': [], 'folders': {}}}
                
                for root, dirs, files in os.walk(temp_dir):
                    rel_path = os.path.relpath(root, temp_dir)
                    if rel_path == '.':
                        rel_path = ''
                    
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in word_exts:
                            full_path = os.path.join(root, file)
                            file_rel_path = os.path.relpath(full_path, temp_dir)
                            
                            file_info = {
                                'name': file,
                                'path': full_path,
                                'rel_path': file_rel_path,
                                'size': os.path.getsize(full_path),
                            }
                            all_files.append(file_info)
                            
                            # 构建文件夹树
                            folder_parts = os.path.dirname(file_rel_path).split(os.sep) if os.path.dirname(file_rel_path) else []
                            
                            current = folder_tree['_root']['folders']
                            for part in folder_parts:
                                if part:
                                    if part not in current:
                                        current[part] = {'files': [], 'folders': {}}
                                    current = current[part]['folders']
                            
                            # 将文件添加到对应层级的 files 列表
                            if folder_parts and folder_parts[0]:
                                target = folder_tree['_root']['folders']
                                for part in folder_parts[:-1]:
                                    target = target[part]['folders']
                                last_folder = folder_parts[-1]
                                if last_folder in target:
                                    target[last_folder]['files'].append(file_info)
                            else:
                                folder_tree['_root']['files'].append(file_info)
                
                st.session_state.uploaded_files = all_files
                st.session_state.folder_tree = folder_tree['_root']
                
                for f in all_files:
                    st.session_state.file_selection[f['path']] = False
                
                st.success(f"✅ 找到 {len(all_files)} 个 Word 文件")
                
            except Exception as e:
                st.error(f"❌ 解压失败：{str(e)}")
                return
    
    # ========== 步骤 2: 选择文件 ==========
    if st.session_state.uploaded_files:
        st.markdown("### ✅ 步骤 2: 选择要转换的文件")
        
        selected_count = sum(st.session_state.file_selection.values())
        total_count = len(st.session_state.uploaded_files)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📊 总文件数", total_count)
        with col2:
            st.metric("✅ 已选择", selected_count)
        with col3:
            st.metric("⭕ 未选择", total_count - selected_count)
        
        st.markdown("#### 批量操作:")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("✅ 全选", use_container_width=True, key="select_all"):
                for f in st.session_state.uploaded_files:
                    st.session_state.file_selection[f['path']] = True
                st.rerun()
        with col2:
            if st.button("❌ 清空", use_container_width=True, key="clear_all"):
                for f in st.session_state.uploaded_files:
                    st.session_state.file_selection[f['path']] = False
                st.rerun()
        with col3:
            if st.button("�� 反选", use_container_width=True, key="invert"):
                for f in st.session_state.uploaded_files:
                    st.session_state.file_selection[f['path']] = not st.session_state.file_selection[f['path']]
                st.rerun()
        with col4:
            btn_text = "📂 收起全部" if st.session_state.expand_all else "📁 展开全部"
            if st.button(btn_text, use_container_width=True, key="toggle_expand"):
                st.session_state.expand_all = not st.session_state.expand_all
                st.rerun()
        
        st.markdown("---")
        
        display_folder_tree(st.session_state.folder_tree, 0)
        
        st.markdown("---")
        st.markdown("### 🚀 步骤 3: 执行转换")
        
        selected_files = [f for f in st.session_state.uploaded_files 
                         if st.session_state.file_selection.get(f['path'], False)]
        
        if selected_files:
            st.success(f"✔️ 已选择 {len(selected_files)} 个文件准备转换")
        else:
            st.warning("⚠️ 请至少选择一个文件")
        
        if st.button("🚀 开始批量转换 PDF", type="primary", use_container_width=True, key="convert_btn"):
            execute_conversion()
        
        if st.session_state.converted_files:
            display_results()


def display_folder_tree(tree, level=0):
    """递归显示文件夹树"""
    if not tree:
        return
    
    indent = " " * (level * 2)
    
    # 显示当前层的文件
    if 'files' in tree and tree['files']:
        for i, file_info in enumerate(tree['files']):
            is_selected = st.session_state.file_selection.get(file_info['path'], False)
            # 使用唯一且稳定的 key
            checkbox_key = f"file_{i}_{file_info['path'].replace(os.sep, '_').replace(' ', '_')}"
            
            col1, col2 = st.columns([10 + level, 2])
            with col1:
                # 直接使用 session state 的值作为 checkbox 的 value
                new_val = st.checkbox(
                    f"{indent}📄 {file_info['name']}",
                    value=is_selected,
                    key=checkbox_key
                )
                # 立即更新 session state
                st.session_state.file_selection[file_info['path']] = new_val
            with col2:
                size_str = format_size(file_info['size'])
                st.caption(f"{size_str}")
    
    # 显示子文件夹
    if 'folders' in tree and tree['folders']:
        for folder_name, folder_data in tree['folders'].items():
            default_expand = st.session_state.expand_all or level < 1
            file_count = len(folder_data.get('files', [])) if isinstance(folder_data, dict) else 0
            with st.expander(f"📁 {indent}{folder_name} ({file_count}个文件)", 
                           expanded=default_expand):
                if isinstance(folder_data, dict):
                    display_folder_tree(folder_data, level + 1)


def format_size(size_bytes):
    """格式化文件大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def execute_conversion():
    """执行转换，保持文件夹结构"""
    selected_files = [f for f in st.session_state.uploaded_files 
                     if st.session_state.file_selection.get(f['path'], False)]
    
    if not selected_files:
        st.warning("请先选择要转换的文件")
        return
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    temp_output = tempfile.mkdtemp(prefix="pdf_output_")
    converted_files = []
    
    try:
        total = len(selected_files)
        status_text.text(f"准备转换 {total} 个文件...")
        
        for i, file_info in enumerate(selected_files):
            file_path = file_info['path']
            original_name = file_info['name']
            rel_path = file_info.get('rel_path', original_name)
            
            output_rel_path = os.path.splitext(rel_path)[0] + '.pdf'
            output_dir = os.path.dirname(os.path.join(temp_output, output_rel_path))
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(temp_output, output_rel_path)
            
            status_text.text(f"转换中：{i+1}/{total} - {original_name}")
            progress_bar.progress((i + 1) / total)
            
            success, error = convert_single_file(file_path, output_path)
            
            if success and os.path.exists(output_path):
                with open(output_path, 'rb') as f:
                    pdf_data = f.read()
                converted_files.append({
                    'name': os.path.basename(output_rel_path),
                    'original_name': original_name,
                    'rel_path': output_rel_path,
                    'success': True,
                    'data': pdf_data
                })
            else:
                converted_files.append({
                    'name': original_name,
                    'original_name': original_name,
                    'success': False,
                    'error': error or '转换失败'
                })
        
        st.session_state.converted_files = converted_files
        status_text.text("✅ 转换完成!")
        
    except Exception as e:
        st.error(f"❌ 转换过程中发生错误：{str(e)}")
    finally:
        try:
            shutil.rmtree(temp_output, ignore_errors=True)
        except:
            pass
    
    st.rerun()


def display_results():
    """显示转换结果"""
    st.markdown("### 📊 转换结果统计")
    
    converted_files = st.session_state.converted_files
    success_count = sum(1 for f in converted_files if f['success'])
    total_count = len(converted_files)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📋 总计", total_count)
    with col2:
        st.metric("✅ 成功", success_count)
    with col3:
        st.metric("❌ 失败", total_count - success_count)
    with col4:
        rate = f"{(success_count/total_count*100):.1f}%" if total_count > 0 else "0%"
        st.metric("📈 成功率", rate)
    
    if success_count > 0:
        st.markdown("---")
        st.markdown("### 📥 下载转换结果")
        
        pdf_files = [f for f in converted_files if f['success']]
        
        if len(pdf_files) == 1:
            pdf_data = pdf_files[0]['data']
            pdf_name = os.path.basename(pdf_files[0]['rel_path']) if 'rel_path' in pdf_files[0] else pdf_files[0]['name']
            st.download_button(
                label=f"📥 下载 PDF 文件 ({format_size(len(pdf_data))})",
                data=pdf_data,
                file_name=pdf_name,
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        else:
            zip_data = create_zip_with_structure(pdf_files)
            zip_name = f"converted_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            st.download_button(
                label=f"📥 下载全部 PDF (ZIP, {len(pdf_files)}个文件，{format_size(len(zip_data))})",
                data=zip_data,
                file_name=zip_name,
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
    
    failed_files = [f for f in converted_files if not f['success']]
    if failed_files:
        with st.expander("❌ 查看失败详情", expanded=True):
            for f in failed_files:
                st.error(f"**{f['original_name']}**: {f.get('error', '未知错误')}")
    
    success_files = [f for f in converted_files if f['success']]
    if success_files:
        with st.expander("✅ 查看成功详情", expanded=False):
            for f in success_files:
                rel_path_info = f.get('rel_path', f['name'])
                st.success(f"`{f['original_name']}`  `{rel_path_info}`")


def create_zip_with_structure(files):
    """创建 ZIP 文件，保持文件夹结构"""
    import io
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in files:
            archive_name = f['rel_path'] if 'rel_path' in f and f['rel_path'] else f['name']
            zipf.writestr(archive_name, f['data'])
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


render()
