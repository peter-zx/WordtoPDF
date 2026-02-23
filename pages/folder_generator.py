"""
文件夹生成页面 - 简单卡片布局
"""

import streamlit as st
import os
import services.folder_service as folder_service
import services.file_service as file_service
import components.folder_selector as folder_selector


def render():
    """渲染文件夹生成页面"""

    # ==================== 步骤1: 输入文件夹名称 ====================
    st.markdown("## 步骤1: 输入文件夹名称")

    # 卡片容器
    st.markdown("""
    <div style="background: #e3f2fd; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
    """, unsafe_allow_html=True)

    # 交互体1: 选择文件
    st.markdown("### 📄 选择文件 (XLSX/TXT/CSV)")

    uploaded_file = st.file_uploader(
        "选择文件",
        type=['xlsx', 'txt', 'csv']
    )

    if uploaded_file:
        temp_dir = os.path.join(os.getcwd(), "temp_input")
        os.makedirs(temp_dir, exist_ok=True)

        temp_file = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_file, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        ext = uploaded_file.name.split('.')[-1].lower()

        try:
            if ext == 'xlsx':
                folder_names = folder_service.FolderService.read_excel_file(temp_file)
            elif ext == 'csv':
                folder_names = folder_service.FolderService.read_csv_file(temp_file)
            else:
                folder_names = folder_service.FolderService.read_text_file(temp_file)

            st.success(f"✅ 成功读取 {len(folder_names)} 个文件夹名称")
            st.session_state.folder_names = folder_names

        except Exception as e:
            st.error(f"❌ 读取文件失败: {str(e)}")

    # 分割线
    st.markdown("---")

    # 交互体2: 粘贴文件夹名称
    st.markdown("### 📝 粘贴文件夹名称")

    text_input = st.text_area(
        "粘贴文件夹名称",
        placeholder="每行一个文件夹路径\n支持多级结构,例如:\n工具 - 钓鱼竿工具 - 锤子工具 - 剃刀工具\n生存建筑 - 烹饪锅 - 冰箱 - 鸟笼",
        height=200,
        key="fg_text_input"
    )

    if text_input:
        folder_names = folder_service.FolderService.parse_text_to_names(text_input)

        if folder_names:
            st.success(f"✅ 解析到 {len(folder_names)} 个文件夹路径")
            st.session_state.folder_names = folder_names

    st.markdown("</div>", unsafe_allow_html=True)

    # ==================== 步骤2: 选择复制文件 ====================
    st.markdown("## 步骤2: 选择复制文件")

    # 卡片容器
    st.markdown("""
    <div style="background: #bbdefb; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
    """, unsafe_allow_html=True)

    # 交互体: 选择需要复制文件
    st.markdown("### 📂 选择需要复制文件")

    uploaded_files = st.file_uploader(
        "选择文件",
        accept_multiple_files=True
    )

    if uploaded_files:
        temp_dir = os.path.join(os.getcwd(), "temp_templates")
        selected_files = file_service.FileService.save_uploaded_files(
            uploaded_files,
            temp_dir
        )

        st.success(f"✅ 已选择 {len(selected_files)} 个文件")

        # 文件选择器(勾选框)
        selected_indices = []
        for i, file_path in enumerate(selected_files):
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            size_str = file_service.FileService.format_file_size(file_size)

            if st.checkbox(f"{file_name} ({size_str})", key=f"file_{i}", value=True):
                selected_indices.append(i)

        selected_files = [selected_files[i] for i in selected_indices]
        st.session_state.template_files = selected_files

    st.markdown("</div>", unsafe_allow_html=True)

    # ==================== 步骤3: 开始生成 ====================
    st.markdown("## 步骤3: 开始生成")

    # 卡片容器
    st.markdown("""
    <div style="background: #90caf9; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
    """, unsafe_allow_html=True)

    # 交互体1: 选择导出路径
    st.markdown("### 🎯 选择导出路径")

    # 默认桌面路径
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    target_dir = folder_selector.render_folder_selector(
        "路径",
        desktop_path
    )

    use_timestamp = st.checkbox(
        "创建带时间戳的顶层文件夹",
        value=True
    )

    # 分割线
    st.markdown("---")

    # 交互体2: 开始生成
    st.markdown("### 🚀 开始生成")

    can_execute = (
        st.session_state.get('folder_names', []) and
        st.session_state.get('template_files', []) and
        target_dir
    )

    if not can_execute:
        if not st.session_state.get('folder_names', []):
            st.warning("⚠️ 请先输入文件夹名称")
        if not st.session_state.get('template_files', []):
            st.warning("⚠️ 请先选择复制文件")
        if not target_dir:
            st.warning("⚠️ 请指定导出路径")

    if st.button("🚀 开始生成", disabled=not can_execute, type="primary", use_container_width=True):
        execute_generation(target_dir, use_timestamp)

    # 显示结果
    if st.session_state.get('generation_result'):
        result = st.session_state.generation_result

        if 'error' in result:
            st.error(f"❌ {result['error']}")
        else:
            st.success(f"✅ 操作完成!")
            st.write(f"**成功:** {len(result['success'])} 个")
            st.write(f"**跳过:** {len(result['skipped'])} 个")
            st.write(f"**失败:** {len(result['failed'])} 个")

    st.markdown("</div>", unsafe_allow_html=True)


def execute_generation(target_dir: str, use_timestamp: bool):
    """执行文件夹生成"""
    with st.spinner("正在生成文件夹..."):
        result = folder_service.FolderService.create_folders_with_files(
            folder_names=st.session_state.folder_names,
            target_dir=target_dir,
            template_files=st.session_state.template_files,
            use_timestamp=use_timestamp
        )

        st.session_state.generation_result = result

    file_service.FileService.cleanup_temp_directory(os.path.join(os.getcwd(), "temp_input"))
    file_service.FileService.cleanup_temp_directory(os.path.join(os.getcwd(), "temp_templates"))

    st.rerun()
