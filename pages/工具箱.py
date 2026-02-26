"""
工具箱页面 - 卡片布局
"""

import streamlit as st

st.title("🛠️ 工具箱")
st.markdown("---")

# 注入卡片样式
st.markdown("""
<style>
    .tool-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 16px;
        padding: 30px 20px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        transition: all 0.3s ease;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .tool-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    .tool-card.disabled {
        background: linear-gradient(135deg, #e0e0e0 0%, #bdbdbd 100%);
        box-shadow: none;
        cursor: not-allowed;
    }
    .tool-card.disabled:hover {
        transform: none;
    }
    .tool-card-icon {
        font-size: 36px;
        margin-bottom: 12px;
    }
    .tool-card-title {
        font-size: 20px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 1px;
    }
    .tool-card-desc {
        font-size: 12px;
        opacity: 0.85;
        margin-top: 8px;
    }
    div[data-testid="stVerticalBlock"] > div[style*="flex-direction: column"] > div > div > div {
        padding: 8px;
    }
</style>
""", unsafe_allow_html=True)

# 工具列表
tools = [
    {"name": "批量PDF转换", "icon": "📑", "desc": "DOCX转PDF", "page": "pages/批量PDF转换.py", "enabled": True},
    {"name": "文档格式转换", "icon": "📄", "desc": "多格式转换", "page": "pages/文档格式转换.py", "enabled": True},
    {"name": "文件重命名", "icon": "🔧", "desc": "待开发", "page": None, "enabled": False},
    {"name": "图片处理", "icon": "🖼️", "desc": "待开发", "page": None, "enabled": False},
    {"name": "文本提取", "icon": "📝", "desc": "待开发", "page": None, "enabled": False},
    {"name": "批量压缩", "icon": "📦", "desc": "待开发", "page": None, "enabled": False},
    {"name": "文件合并", "icon": "🔗", "desc": "待开发", "page": None, "enabled": False},
    {"name": "格式转换", "icon": "🔄", "desc": "待开发", "page": None, "enabled": False},
    {"name": "文件夹生成", "icon": "📂", "desc": "待开发", "page": None, "enabled": False},
    {"name": "批量水印", "icon": "💧", "desc": "待开发", "page": None, "enabled": False},
    {"name": "文本替换", "icon": "✏️", "desc": "待开发", "page": None, "enabled": False},
    {"name": "更多工具", "icon": "➕", "desc": "敬请期待", "page": None, "enabled": False},
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
                    st.markdown(f"""
                    <div class="tool-card" onclick="Streamlit.setComponentValue({{'value': 'tool_{idx}_click'}})">
                        <div class="tool-card-icon">{tool['icon']}</div>
                        <div class="tool-card-title">{tool['name']}</div>
                        <div class="tool-card-desc">{tool['desc']}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 添加隐藏按钮处理点击
                    if st.button("", key=f"tool_{idx}_click"):
                        st.switch_page(tool["page"])
                else:
                    st.markdown(f"""
                    <div class="tool-card disabled">
                        <div class="tool-card-icon">{tool['icon']}</div>
                        <div class="tool-card-title">{tool['name']}</div>
                        <div class="tool-card-desc">{tool['desc']}</div>
                    </div>
                    """, unsafe_allow_html=True)
