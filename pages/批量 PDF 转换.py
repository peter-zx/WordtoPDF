"""
批量 PDF 转换独立应用 - 完整版本
功能:
1. 上传 ZIP 压缩包，解析文件夹结构
2. 树形显示文件 (支持展开/折叠)
3. 多选文件 (全选/清空/反选),实时计数
4. 保持文件夹结构批量转换为 PDF
5. 打包下载 (保持原始结构)
"""

import os
import streamlit as st
import tempfile
import zipfile
import shutil
from datetime import datetime

# 导入转换服务
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from services.word_converter import convert_single_file


def main():
    """主函数"""
    st.set_page_config(
        page_title="批量 PDF 转换",
        page_icon="📑",
        layout="wide"
    )
    
    # 自定义 CSS
    st.markdown("""
    <style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
    }
    .folder-tree-item {
        margin-left: 20px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📑 批量 PDF 转换工具")
    st.markdown("---")
    
    # 初始化 session state
    init_session_state()
    
    # ========== 步骤 1: 上传 ZIP 压缩包 ==========
    render_upload_section()
    
    # ========== 步骤 2: 选择文件 ==========
    if st.session_state.uploaded_files:
        render_selection_section()
        
        # ========== 步骤 3: 执行转换 ==========
        render_conversion_section()
        
        # ========== 显示结果 ==========
        if st.session_state.converted_files:
            render_results_section()


def init_session_state():
    """初始化 session state"""
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = []
    if 'file_selection' not in st.session_state:
        st.session_state.file_selection = {}
    if 'folder_selection' not in st.session_state:
        st.session_state.folder_selection = {}
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = []
    if 'folder_tree' not in st.session_state:
        st.session_state.folder_tree = {'files': [], 'folders': {}}
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = None
    if 'expand_all' not in st.session_state:
        st.session_state.expand_all = False
    if 'conversion_started' not in st.session_state:
        st.session_state.conversion_started = False


def render_upload_section():
    """渲染上传区域"""
    st.markdown("### 📂 步骤 1: 上传文件夹 (ZIP 压缩包)")
    
    st.info("""
    💡 **使用说明**:
    1. 将包含 Word 文件的整个文件夹压缩为 **ZIP 格式**
    2. 系统会保持原始文件夹结构显示所有文件
    3. 支持多层嵌套文件夹 (如：部门/员工/合同.docx)
    """)
    
    uploaded_zip = st.file_uploader(
        "点击选择 ZIP 文件上传",
        type=['zip'],
        help="上传包含 DOC/DOCX 文件的文件夹压缩包",
        key="zip_uploader"
    )
    
    if uploaded_zip:
        st.success(f"✅ 已上传：{uploaded_zip.name} ({uploaded_zip.size / 1024:.1f} KB)")
        process_uploaded_zip(uploaded_zip)


def process_uploaded_zip(uploaded_zip):
    """处理上传的 ZIP 文件"""
    with st.spinner("📦 正在解压并分析文件夹结构..."):
        try:
            # 清理旧临时目录
            if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
                shutil.rmtree(st.session_state.temp_dir, ignore_errors=True)
            
            # 创建新临时目录并解压
            temp_dir = tempfile.mkdtemp(prefix="pdf_input_")
            st.session_state.temp_dir = temp_dir
            
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                # 处理中文编码问题
                for info in zip_ref.infolist():
                    # 尝试解码中文文件名
                    try:
                        # GBK 编码 (中文 Windows)
                        info.filename = info.filename.encode('cp437').decode('gbk')
                    except:
                        pass
                    zip_ref.extract(info, temp_dir)
            
            # 扫描所有 Word 文件并构建树形结构
            word_exts = ['.doc', '.docx']
            all_files = []
            folder_tree = {'files': [], 'folders': {}}
            
            # 调试：打印所有找到的文件
            debug_folders = set()
            
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
                        
                        # 记录调试信息
                        folder_of_file = os.path.dirname(file_rel_path)
                        if folder_of_file:
                            top_level_folder = folder_of_file.split(os.sep)[0]
                            debug_folders.add(top_level_folder)
                        
                        # 构建文件夹树结构
                        add_file_to_tree(folder_tree, file_info)
            
            # 调试信息
            print(f"✅ 找到 {len(all_files)} 个文件")
            print(f"📁 第一层文件夹：{sorted(debug_folders)}")
            
            # 保存到 session state
            st.session_state.uploaded_files = all_files
            st.session_state.folder_tree = folder_tree
            
            # 初始化所有文件和文件夹为未选中
            for f in all_files:
                st.session_state.file_selection[f['path']] = False
            
            # 初始化文件夹选择状态
            init_folder_selection(folder_tree)
            
            st.success(f"✅ 找到 {len(all_files)} 个 Word 文件，共 {len(debug_folders)} 个一级文件夹")
            
        except Exception as e:
            st.error(f"❌ 解压失败：{str(e)}")
            import traceback
            st.error(traceback.format_exc())


def init_folder_selection(tree, path=''):
    """初始化文件夹选择状态"""
    if not tree:
        return
    
    # 当前路径作为 key
    if path:
        st.session_state.folder_selection[path] = False
    
    # 递归初始化子文件夹
    for folder_name, folder_data in tree.get('folders', {}).items():
        new_path = f"{path}/{folder_name}" if path else folder_name
        init_folder_selection(folder_data, new_path)


def add_file_to_tree(tree, file_info):
    """将文件添加到树形结构中"""
    folder_path = os.path.dirname(file_info['rel_path'])
    
    if not folder_path:
        # 根目录文件
        tree['files'].append(file_info)
    else:
        # 嵌套文件夹文件 - 使用完整路径作为唯一标识
        parts = folder_path.split(os.sep)
        current = tree['folders']
        
        for i, part in enumerate(parts):
            if part:
                if part not in current:
                    current[part] = {'files': [], 'folders': {}}
                
                # 如果是最后一层，添加文件
                if i == len(parts) - 1:
                    current[part]['files'].append(file_info)
                
                current = current[part]['folders']


def render_selection_section():
    """渲染文件选择区域 - 简化版本"""
    st.markdown("### ✅ 步骤 2: 选择要转换的文件")
    
    # 上传后默认全选所有文件
    if st.session_state.uploaded_files and not st.session_state.conversion_started:
        for f in st.session_state.uploaded_files:
            st.session_state.file_selection[f['path']] = True
    
    st.info("✅ 已自动选中所有上传的文件，您可以直接在下方点击「开始批量转换 PDF」", icon="ℹ️")
    
    # 显示文件夹树 (Windows 风格) - 只读展示
    with st.expander("📁 查看/调整文件选择 (可选)", expanded=False):
        display_folder_tree_windows(st.session_state.folder_tree, 0, '')


def select_all_files(select):
    """全选或清空所有文件"""
    for f in st.session_state.uploaded_files:
        st.session_state.file_selection[f['path']] = select
    
    # 同时设置文件夹状态
    select_all_folders(select)


def select_all_folders(select):
    """全选或清空所有文件夹"""
    def set_folder_state(tree, path=''):
        if path:
            st.session_state.folder_selection[path] = select
        for folder_name, folder_data in tree.get('folders', {}).items():
            new_path = f"{path}/{folder_name}" if path else folder_name
            set_folder_state(folder_data, new_path)
    
    set_folder_state(st.session_state.folder_tree)


def invert_selection():
    """反选所有文件"""
    for f in st.session_state.uploaded_files:
        current = st.session_state.file_selection.get(f['path'], False)
        st.session_state.file_selection[f['path']] = not current


def display_folder_tree_windows(tree, level=0, path=''):
    """Windows 风格递归显示文件夹树，勾选框在左侧"""
    import hashlib
    
    if not tree:
        return
    
    # 显示当前层的文件
    if tree.get('files'):
        for i, file_info in enumerate(tree['files']):
            is_selected = st.session_state.file_selection.get(file_info['path'], False)
            unique_id = hashlib.md5(file_info['path'].encode('utf-8')).hexdigest()[:12]
            checkbox_key = f"file_{unique_id}"
            
            # 使用容器包裹，紧凑布局
            with st.container():
                # 文件：勾选框 + 文件名
                new_val = st.checkbox(
                    f"{'  ' * level}📄 {file_info['name']}",
                    value=is_selected,
                    key=checkbox_key
                )
                # 立即更新状态
                if new_val != is_selected:
                    st.session_state.file_selection[file_info['path']] = new_val
    
    # 显示子文件夹 - Windows 风格
    if tree.get('folders'):
        for folder_name, folder_data in sorted(tree['folders'].items()):
            folder_path = f"{path}/{folder_name}" if path else folder_name
            
            # 获取文件夹内所有文件
            folder_files = get_all_files_in_folder(folder_data)
            selected_count = sum(1 for f in folder_files 
                                if st.session_state.file_selection.get(f['path'], False))
            total_count = len(folder_files)
            
            # 文件夹选择状态
            is_folder_selected = st.session_state.folder_selection.get(folder_path, False)
            folder_checkbox_key = f"folder_{folder_path.replace('/', '_')}"
            
            # 计算文件夹状态图标
            if total_count > 0:
                if selected_count == total_count:
                    folder_icon = "✅"  # 全选
                elif selected_count > 0:
                    folder_icon = "⭕"  # 部分选中
                else:
                    folder_icon = "📁"  # 未选中
            else:
                folder_icon = "📁"
            
            # 默认展开级别
            default_expand = st.session_state.expand_all or level < 1
            
            # 文件夹标题 (带展开/折叠)
            folder_label = f"{'  ' * level}{folder_icon} {folder_name} ({selected_count}/{total_count})"
            
            with st.expander(folder_label, expanded=default_expand):
                # 文件夹勾选框 (在 expander 内部第一行)
                folder_sel = st.checkbox(
                    f"选中 '{folder_name}' 下所有 {total_count} 个文件",
                    value=is_folder_selected,
                    key=folder_checkbox_key,
                    label_visibility="collapsed"
                )
                
                # 处理文件夹勾选逻辑 - 不刷新，让 Streamlit 自动处理
                if folder_sel != is_folder_selected:
                    st.session_state.folder_selection[folder_path] = folder_sel
                    
                    if folder_sel:
                        # 选中该文件夹下所有文件
                        for f in folder_files:
                            st.session_state.file_selection[f['path']] = True
                    else:
                        # 取消勾选时清空该文件夹下的选择
                        for f in folder_files:
                            st.session_state.file_selection[f['path']] = False
                
                # 递归显示子内容
                display_folder_tree_windows(folder_data, level + 1, folder_path)


def get_all_files_in_folder(tree):
    """获取文件夹内所有文件 (包括子文件夹)"""
    files = []
    if not tree:
        return files
    
    files.extend(tree.get('files', []))
    
    for folder_data in tree.get('folders', {}).values():
        files.extend(get_all_files_in_folder(folder_data))
    
    return files


def format_size(size_bytes):
    """格式化文件大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def render_conversion_section():
    """渲染转换区域"""
    st.markdown("---")
    st.markdown("### 🚀 步骤 3: 执行转换")
    
    selected_files = [
        f for f in st.session_state.uploaded_files 
        if st.session_state.file_selection.get(f['path'], False)
    ]
    
    if selected_files:
        # 使用卡片式布局，醒目的数字显示
        st.markdown("""
        <style>
        .conversion-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 30px;
            border-radius: 15px;
            text-align: center;
            margin: 20px 0;
            box-shadow: 0 4px 15px rgba(102,126,234,0.3);
        }
        .conversion-number {
            font-size: 64px;
            font-weight: bold;
            color: white;
            margin: 10px 0;
        }
        .conversion-label {
            font-size: 18px;
            color: white;
            opacity: 0.9;
        }
        </style>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="conversion-card">
            <div class="conversion-label">✅ 已选择</div>
            <div class="conversion-number">{len(selected_files)}</div>
            <div class="conversion-label">个文件准备转换</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.warning("⚠️ 请至少选择一个文件", icon="⚠️")
    
    if st.button("🚀 开始批量转换 PDF", type="primary", use_container_width=True, key="convert_btn"):
        execute_conversion(selected_files)


def execute_conversion(selected_files):
    """执行转换"""
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
            
            # 保持文件夹结构生成输出路径
            output_rel_path = os.path.splitext(rel_path)[0] + '.pdf'
            output_dir = os.path.dirname(os.path.join(temp_output, output_rel_path))
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(temp_output, output_rel_path)
            
            # 计算百分比
            percentage = int((i + 1) / total * 100)
            status_text.text(f"转换中：{i+1}/{total} ({percentage}%) - {original_name}")
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
        st.rerun()
        
    except Exception as e:
        st.error(f"❌ 转换过程中发生错误：{str(e)}")
    finally:
        try:
            shutil.rmtree(temp_output, ignore_errors=True)
        except:
            pass


def render_results_section():
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
    
    # 显示失败详情
    failed_files = [f for f in converted_files if not f['success']]
    if failed_files:
        with st.expander("❌ 查看失败详情", expanded=True):
            for f in failed_files:
                st.error(f"**{f['original_name']}**: {f.get('error', '未知错误')}")
    
    # 显示成功详情
    success_files = [f for f in converted_files if f['success']]
    if success_files:
        with st.expander("✅ 查看成功详情", expanded=False):
            for f in success_files:
                rel_path_info = f.get('rel_path', f['name'])
                st.success(f"`{f['original_name']}` → `{rel_path_info}`")


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


if __name__ == "__main__":
    main()
