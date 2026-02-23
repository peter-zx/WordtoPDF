"""
文件夹选择器组件
"""

import streamlit as st


def render_folder_selector(label: str, default_path: str = ""):
    """
    渲染文件夹选择器

    Args:
        label: 标签
        default_path: 默认路径

    Returns:
        选择的文件夹路径
    """
    st.markdown(f"""
    <div style="display: flex; gap: 8px; align-items: center; margin-bottom: 16px;">
        <label style="font-weight: 600; min-width: 120px;">{label}</label>
        <input type="text" id="folder_{label}" value="{default_path}" 
               style="flex: 1; padding: 8px 12px; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 0.9rem;">
        <button onclick="selectFolder('folder_{label}')" 
                style="padding: 8px 16px; background: #4299e1; color: white; border: none; border-radius: 6px; cursor: pointer;">
            📁 浏览
        </button>
    </div>
    <script>
    function selectFolder(inputId) {{
        // 由于浏览器安全限制,Web应用无法直接访问文件系统
        // 这里只是模拟,实际需要使用桌面应用或后端API
        alert('请在输入框中手动输入文件夹路径');
    }}
    </script>
    """, unsafe_allow_html=True)

    # 由于Streamlit限制,我们还是使用text_input,但添加提示
    path = st.text_input(
        label,
        value=default_path,
        help="💡 请输入文件夹路径,或从文件管理器复制路径"
    )

    return path


def render_file_selector(label: str, default_path: str = ""):
    """
    渲染文件选择器

    Args:
        label: 标签
        default_path: 默认路径

    Returns:
        选择的文件路径
    """
    path = st.text_input(
        label,
        value=default_path,
        help="💡 请输入文件路径,或从文件管理器复制路径"
    )

    return path
