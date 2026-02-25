"""
文件批量处理工具箱 - Streamlit多页面版本
"""

import streamlit as st
import config

# 页面配置
st.set_page_config(**config.PAGE_CONFIG)

# 初始化session state
if 'folder_names' not in st.session_state:
    st.session_state.folder_names = []
if 'template_files' not in st.session_state:
    st.session_state.template_files = []
if 'generation_result' not in st.session_state:
    st.session_state.generation_result = None

# 自定义CSS隐藏默认导航,使用中文导航
st.markdown("""
<style>
    /* 隐藏默认的侧边栏导航 */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 自定义导航样式 */
    .custom-nav {
        padding: 1rem 0;
    }

    .custom-nav-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #0d47a1;
        margin-bottom: 1rem;
        padding: 0.5rem;
        border-bottom: 2px solid #e3f2fd;
    }

    .nav-section {
        margin-bottom: 1.5rem;
    }

    .nav-section-title {
        font-size: 0.9rem;
        font-weight: 600;
        color: #718096;
        margin-bottom: 0.5rem;
        padding-left: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# 在侧边栏创建自定义导航
with st.sidebar:
    st.markdown('<div class="custom-nav">', unsafe_allow_html=True)
    st.markdown('<div class="custom-nav-title">📂 导航菜单</div>', unsafe_allow_html=True)

    # 主页
    if st.button("📁 文件夹生成", use_container_width=True, type="primary"):
        st.switch_page("pages/文件夹生成.py")

    # 工具箱
    if st.button("🛠️ 工具箱", use_container_width=True):
        st.switch_page("pages/工具箱.py")

    st.markdown('</div>', unsafe_allow_html=True)

# 主页内容
st.title("📁 文件夹批量生成工具")
st.markdown("---")

# 导入并渲染文件夹生成页面
import pages.folder_generator as folder_generator
folder_generator.render()
