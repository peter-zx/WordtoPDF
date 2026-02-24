"""
工具箱页面
"""

import streamlit as st

st.title("🛠️ 工具箱")
st.markdown("---")

st.markdown("### 可用工具")

# 工具卡片
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div style="background: #e3f2fd; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
        <h3>📄 文档格式转换</h3>
        <p>将RTF、ODT、HTML等格式转换为DOCX格式</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("使用工具", key="tool1", use_container_width=True):
        st.switch_page("pages/文档格式转换.py")

with col2:
    st.markdown("""
    <div style="background: #bbdefb; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
        <h3>📑 批量PDF转换</h3>
        <p>批量将DOCX文件转换为PDF格式</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("使用工具", key="tool2", use_container_width=True):
        st.info("该功能正在开发中...")

col3, col4 = st.columns(2)

with col3:
    st.markdown("""
    <div style="background: #90caf9; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
        <h3>🔧 文件重命名</h3>
        <p>批量重命名文件</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("使用工具", key="tool3", use_container_width=True):
        st.info("该功能正在开发中...")

with col4:
    st.markdown("""
    <div style="background: #64b5f6; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
        <h3>⚡ 更多工具</h3>
        <p>更多实用工具即将推出</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("敬请期待", key="tool4", use_container_width=True):
        st.info("更多工具正在开发中...")
