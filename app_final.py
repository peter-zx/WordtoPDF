"""
文件批量处理工具箱 - 最终版本
模块化架构,两个页面
"""

import streamlit as st
import os
import config
import styles_dashboard
import components.sidebar as sidebar
import pages.folder_generator as folder_generator
import pages.toolbox as toolbox


def init_session_state():
    """初始化session state"""
    if 'current_page' not in st.session_state:
        st.session_state.current_page = 'folder_generator'

    # 文件夹生成页面状态
    if 'folder_names' not in st.session_state:
        st.session_state.folder_names = []
    if 'template_files' not in st.session_state:
        st.session_state.template_files = []
    if 'generation_result' not in st.session_state:
        st.session_state.generation_result = None


def render_top_nav():
    """渲染顶部导航栏"""
    st.markdown(f"""
    <div class="top-nav">
        <h1>{config.PAGE_CONFIG['page_title']}</h1>
        <div class="top-nav-subtitle">{config.PAGE_CONFIG.get('page_subtitle', '')}</div>
    </div>
    """, unsafe_allow_html=True)


def main():
    """主函数"""
    # 页面配置
    st.set_page_config(**config.PAGE_CONFIG)

    # 应用样式
    st.markdown(styles_dashboard.STYLES_DASHBOARD, unsafe_allow_html=True)

    # 初始化session state
    init_session_state()

    # 渲染顶部导航栏
    render_top_nav()

    # 创建分栏布局
    col_sidebar, col_main = st.columns([1, 4])

    with col_sidebar:
        # 渲染侧边栏
        nav_items = [
            {
                'key': 'folder_generator',
                'icon': '📁',
                'label': '文件夹生成'
            },
            {
                'key': 'toolbox',
                'icon': '🛠️',
                'label': '工具箱'
            }
        ]

        sidebar.Sidebar.render(nav_items, st.session_state.current_page)

    with col_main:
        # 渲染主内容区
        st.markdown('<div class="main-content">', unsafe_allow_html=True)

        if st.session_state.current_page == 'folder_generator':
            folder_generator.render()
        elif st.session_state.current_page == 'toolbox':
            toolbox.render()

        st.markdown('</div>', unsafe_allow_html=True)

    # 页脚
    st.markdown("""
    <div style='text-align: center; color: #718096; font-size: 0.85rem; padding: 24px; margin-top: 24px;'>
        💡 所有操作在本地执行,数据不会上传到任何服务器
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
