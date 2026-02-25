"""
工具箱页面
"""

import streamlit as st

st.title("🛠️ 工具箱")
st.markdown("---")

st.markdown("### 可用工具列表")

# 文档格式转换 - 大按钮卡片
st.markdown("""
<style>
div.stButton > button {
    min-height: 180px !important;
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    padding: 40px !important;
    line-height: 1.6 !important;
}
</style>
""", unsafe_allow_html=True)

if st.button("📄 文档格式转换", key="tool1", type="primary", use_container_width=True):
    st.switch_page("pages/文档格式转换.py")

st.markdown("""
<div style="text-align: center; margin-top: -10px; margin-bottom: 30px;">
    <p style="font-size: 1.1rem; color: #424242; margin: 5px 0;">将RTF、ODT、HTML、TXT等格式转换为DOCX格式</p>
    <p style="font-size: 0.95rem; color: #757575; margin: 5px 0;">✅ 支持批量转换 | ✅ 保持文本格式 | ✅ 快速高效</p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# 其他工具 - 预留
st.markdown("### 更多工具")

col1, col2 = st.columns(2)

with col1:
    if st.button("📑 批量PDF转换", key="tool2", use_container_width=True):
        st.switch_page("pages/批量PDF转换.py")
    st.markdown("""
    <div style="text-align: center; margin-top: -10px; margin-bottom: 20px;">
        <p style="font-size: 0.95rem; color: #757575; margin: 5px 0;">批量将DOCX文件转换为PDF格式</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    if st.button("🔧 文件重命名", key="tool3", use_container_width=True, disabled=True):
        pass
    st.markdown("""
    <div style="text-align: center; margin-top: -10px; margin-bottom: 20px; opacity: 0.6;">
        <p style="font-size: 0.95rem; color: #757575; margin: 5px 0;">批量重命名文件</p>
    </div>
    """, unsafe_allow_html=True)
