"""
批量PDF转换页面 - 云端版本（修复版）
使用Streamlit原生状态管理，避免callback问题
"""

import os
import streamlit as st
import tempfile
import zipfile
import shutil
from datetime import datetime
from services.word_converter import convert_single_file


def render():
    """渲染批量PDF转换页面"""
    st.title("📑 批量PDF转换")
    st.markdown("---")
    st.info("💡 使用LibreOffice进行转换。请上传包含Word文件的文件夹(ZIP格式)，转换后下载PDF。")
    
    # 初始化session state
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'file_selection' not in st.session_state:
        st.session_state.file_selection = {}
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []
    if 'show_results' not in st.session_state:
        st.session_state.show_results = False

    # 步骤1: 上传ZIP文件
    st.markdown("## 📂 步骤1: 上传文件夹")
    
    uploaded_zip = st.file_uploader(
        "选择ZIP文件",
        type=['zip'],
        help="上传包含DOC/DOCX文件的文件夹压缩包"
    )

    if uploaded_zip:
        st.success(f"✅ 已上传: {uploaded_zip.name} ({uploaded_zip.size / 1024:.1f} KB)")
        
        # 解压并获取文件列表
        with st.spinner("解压中..."):
            try:
                temp_dir = tempfile.mkdtemp(prefix="pdf_upload_")
                
                # 保存ZIP文件
                zip_path = os.path.join(temp_dir, "upload.zip")
                with open(zip_path, 'wb') as f:
                    f.write(uploaded_zip.getbuffer())
                
                # 解压（处理中文文件名）
                extract_dir = os.path.join(temp_dir, "extracted")
                os.makedirs(extract_dir, exist_ok=True)
                
                with zipfile.ZipFile(zip_path, 'r') as zipf:
                    for member in zipf.namelist():
                        # 跳过目录
                        if member.endswith('/'):
                            continue
                        
                        # 处理文件名编码
                        try:
                            filename = member.encode('cp437').decode('gbk')
                        except:
                            filename = member
                        
                        # 跳过非Word文件
                        if not filename.lower().endswith(('.doc', '.docx')):
                            continue
                        
                        # 解压文件
                        source = zipf.open(member)
                        target_path = os.path.join(extract_dir, filename)
                        os.makedirs(os.path.dirname(target_path), exist_ok=True)
                        
                        with open(target_path, 'wb') as target:
                            target.write(source.read())
                
                # 获取所有Word文件
                word_files = []
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.lower().endswith(('.doc', '.docx')):
                            file_path = os.path.join(root, file)
                            word_files.append({
                                'name': file,
                                'path': file_path,
                                'relative': os.path.relpath(file_path, extract_dir)
                            })
                
                # 初始化选择状态
                st.session_state.uploaded_files = word_files
                st.session_state.temp_dir = temp_dir
                st.session_state.file_selection = {f['path']: False for f in word_files}
                
                if word_files:
                    st.success(f"✅ 找到 {len(word_files)} 个Word文件")
                else:
                    st.warning("⚠️ 未找到Word文件")
                    
            except Exception as e:
                st.error(f"❌ 解压失败: {str(e)}")

    # 显示文件列表和选择
    if st.session_state.uploaded_files:
        st.markdown("---")
        st.markdown("### 📁 文件列表")
        
        # 控制按钮
        col1, col2, col3, col4 = st.columns([1, 1, 1, 1.5])
        
        with col1:
            if st.button("全选", key="btn_all"):
                for f in st.session_state.uploaded_files:
                    st.session_state.file_selection[f['path']] = True
                st.rerun()
        
        with col2:
            if st.button("清空", key="btn_clear"):
                for f in st.session_state.uploaded_files:
                    st.session_state.file_selection[f['path']] = False
                st.rerun()
        
        with col3:
            st.write("")
        
        # 显示已选择计数
        selected_count = sum(1 for selected in st.session_state.file_selection.values() if selected)
        total_count = len(st.session_state.uploaded_files)
        
        with col4:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; padding: 16px 24px; text-align: center; color: white; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);">
                <div style="font-size: 14px; opacity: 0.9;">已选择文件</div>
                <div style="font-size: 32px; font-weight: 700; letter-spacing: 2px;">{selected_count} / {total_count}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # 文件列表（使用checkbox）
        st.markdown("---")
        st.markdown("### 📋 选择要转换的文件")
        
        for i, file_info in enumerate(st.session_state.uploaded_files):
            file_path = file_info['path']
            is_selected = st.session_state.file_selection.get(file_path, False)
            
            # 使用checkbox更新状态
            new_value = st.checkbox(f"📄 {file_info['name']}", value=is_selected, key=f"file_{i}")
            
            # 只有状态改变时才更新
            if new_value != is_selected:
                st.session_state.file_selection[file_path] = new_value

    # 步骤2: 开始转换
    st.markdown("---")
    st.markdown("## 🚀 步骤2: 开始转换")
    
    selected_files = get_selected_files()
    selected_count = len(selected_files)
    can_convert = selected_count > 0
    
    if st.button("开始批量转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion()

    # 显示结果
    if st.session_state.show_results and st.session_state.converted_files:
        st.markdown("---")
        st.markdown("## 📊 转换结果")
        
        converted_files = st.session_state.converted_files
        success_count = sum(1 for f in converted_files if f['success'])
        total_count = len(converted_files)
        
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
            
            pdf_files = [f for f in converted_files if f['success']]
            
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
        failed_files = [f for f in converted_files if not f['success']]
        if failed_files:
            with st.expander("❌ 查看失败详情", expanded=True):
                for f in failed_files:
                    st.error(f"`{f['name']}`: {f.get('error', '未知错误')}")
        
        # 显示成功详情
        success_files = [f for f in converted_files if f['success']]
        if success_files:
            with st.expander("✅ 查看成功详情", expanded=False):
                for f in success_files:
                    st.success(f"`{f['original_name']}` → `{f['name']}`")


def get_selected_files():
    """获取选中的文件列表"""
    return [f for f in st.session_state.uploaded_files 
            if st.session_state.file_selection.get(f['path'], False)]


def execute_conversion():
    """执行转换"""
    selected_files = get_selected_files()
    
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
            file_name = file_info['name']
            output_name = os.path.splitext(file_name)[0] + '.pdf'
            output_path = os.path.join(temp_output, output_name)
            
            status_text.text(f"转换中: {i+1}/{total} - {file_name}")
            progress_bar.progress((i + 1) / total)
            
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
        
        st.session_state.converted_files = converted_files
        st.session_state.show_results = True
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
