"""
工具箱页面 - 卡片布局
"""

import streamlit as st

st.title("🛠️ 工具箱")
st.markdown("---")

# 注入卡片样式
st.markdown("""
<style>
    div[data-testid="stVerticalBlock"] > div[style*="flex-direction: column"] > div > div > div {
        padding: 12px;
    }
    .tool-card-btn {
        height: 120px !important;
        border-radius: 12px !important;
        font-size: 22px !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
        transition: all 0.2s ease !important;
        margin: 0 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }
    .tool-card-btn:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.5) !important;
    }
    .tool-card-btn:disabled {
        opacity: 0.5 !important;
        cursor: not-allowed !important;
        background: linear-gradient(135deg, #e0e0e0 0%, #bdbdbd 100%) !important;
        color: #757575 !important;
    }
</style>
""", unsafe_allow_html=True)

# 工具列表
tools = [
    {"name": "批量 PDF 转换", "page": "apps/批量 PDF 转换.py", "enabled": True},
    {"name": "文档格式转换", "page": "pages/文档格式转换.py", "enabled": True},
    {"name": "文件重命名", "page": None, "enabled": False},
    {"name": "图片处理", "page": None, "enabled": False},
    {"name": "文本提取", "page": None, "enabled": False},
    {"name": "批量压缩", "page": None, "enabled": False},
    {"name": "文件合并", "page": None, "enabled": False},
    {"name": "格式转换", "page": None, "enabled": False},
    {"name": "文件夹生成", "page": None, "enabled": False},
    {"name": "批量水印", "page": None, "enabled": False},
    {"name": "文本替换", "page": None, "enabled": False},
    {"name": "待开发中", "page": None, "enabled": False},
]

# 渲染3行4列卡片
for row in range(3):
    cols = st.columns(4)
    for col_idx in range(4):
        idx = row * 4 + col_idx
        if idx < len(tools):
            tool = tools[idx]
            with cols[col_idx]:
                if tool["enabled"]:
                    if st.button(tool["name"], key=f"tool_{idx}", type="primary", use_container_width=True):
                        st.switch_page(tool["page"])
                else:
                    st.button(tool["name"], key=f"tool_{idx}", disabled=True, use_container_width=True)
