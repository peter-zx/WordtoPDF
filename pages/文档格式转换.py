"""
文档格式转换页面
"""

import streamlit as st
import os
import sys

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import services.format_converter_service as format_converter_service
import services.file_service as file_service


def render():
    """渲染文档格式转换页面"""

    st.title("📄 文档格式转换")
    st.markdown("---")

    # 说明
    st.markdown("""
    ### 功能说明
    将各种文档格式转换为Word文档(DOCX)格式。

    **支持的输入格式:**
    - RTF (Rich Text Format)
    - ODT (OpenDocument Text)
    - HTML (HyperText Markup Language)
    - TXT (Plain Text)
    """)

    st.markdown("---")

    # 步骤1: 选择文件
    st.markdown("## 步骤1: 选择文件")

    uploaded_files = st.file_uploader(
        "选择需要转换的文件",
        type=['rtf', 'odt', 'html', 'htm', 'txt'],
        accept_multiple_files=True
    )

    if uploaded_files:
        # 保存上传的文件
        temp_dir = os.path.join(os.getcwd(), "temp_convert_input")
        os.makedirs(temp_dir, exist_ok=True)

        saved_files = []
        for uploaded_file in uploaded_files:
            temp_file = os.path.join(temp_dir, uploaded_file.name)
            with open(temp_file, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            saved_files.append(temp_file)

        st.success(f"✅ 已选择 {len(saved_files)} 个文件")

        # 显示文件列表
        for file_path in saved_files:
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            size_str = file_service.FileService.format_file_size(file_size)
            st.write(f"📄 {file_name} ({size_str})")

        st.session_state.convert_files = saved_files

    st.markdown("---")

    # 步骤2: 选择输出目录
    st.markdown("## 步骤2: 选择输出目录")

    # 默认桌面路径
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    default_output = os.path.join(desktop_path, "converted_docs")

    output_dir = st.text_input(
        "输出目录路径",
        value=default_output,
        help="转换后的DOCX文件将保存在此目录"
    )

    st.markdown("---")

    # 步骤3: 开始转换
    st.markdown("## 步骤3: 开始转换")

    can_convert = (
        st.session_state.get('convert_files', []) and
        output_dir
    )

    if not can_convert:
        if not st.session_state.get('convert_files', []):
            st.warning("⚠️ 请先选择需要转换的文件")
        if not output_dir:
            st.warning("⚠️ 请指定输出目录")

    if st.button("🚀 开始转换", disabled=not can_convert, type="primary", use_container_width=True):
        execute_conversion(output_dir)

    # 显示结果
    if st.session_state.get('conversion_result'):
        result = st.session_state.conversion_result

        st.markdown("---")
        st.markdown("### 转换结果")

        success_count = sum(1 for r in result if r['success'])
        failed_count = len(result) - success_count

        st.write(f"**成功:** {success_count} 个")
        st.write(f"**失败:** {failed_count} 个")

        # 显示成功列表
        if success_count > 0:
            with st.expander("✅ 查看成功详情", expanded=False):
                for r in result:
                    if r['success']:
                        input_name = os.path.basename(r['input_file'])
                        output_name = os.path.basename(r['output_file'])
                        st.write(f"📄 {input_name} → {output_name}")

        # 显示失败列表
        if failed_count > 0:
            with st.expander("❌ 查看失败详情", expanded=False):
                for r in result:
                    if not r['success']:
                        input_name = os.path.basename(r['input_file'])
                        st.write(f"📄 {input_name}: {r['error']}")


def execute_conversion(output_dir: str):
    """执行转换"""
    with st.spinner("正在转换文件..."):
        results = format_converter_service.FormatConverterService.batch_convert(
            st.session_state.convert_files,
            output_dir
        )

        st.session_state.conversion_result = results

    # 清理临时文件
    temp_dir = os.path.join(os.getcwd(), "temp_convert_input")
    file_service.FileService.cleanup_temp_directory(temp_dir)

    st.success(f"✅ 转换完成!")
    st.rerun()


# Streamlit页面入口
if __name__ == "__main__":
    render()
