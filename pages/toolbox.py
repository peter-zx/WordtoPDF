"""
工具箱页面
"""

import streamlit as st


def render():
    """渲染工具箱页面"""

    st.markdown("### 🛠️ 工具箱")

    st.markdown("各种实用工具集合")

    # 功能卡片网格
    st.markdown('<div class="toolbox-grid">', unsafe_allow_html=True)

    # 工具列表
    tools = [
        {
            'icon': '🔄',
            'name': '格式转换',
            'description': 'RTF/ODT/HTML → DOCX',
            'status': '开发中'
        },
        {
            'icon': '📄',
            'name': 'PDF转换',
            'description': 'DOCX → PDF',
            'status': '开发中'
        },
        {
            'icon': '📊',
            'name': '批量重命名',
            'description': '批量重命名文件',
            'status': '开发中'
        },
        {
            'icon': '🔍',
            'name': '搜索文件',
            'description': '在文件夹中搜索文件',
            'status': '开发中'
        },
        {
            'icon': '📝',
            'name': '批量替换',
            'description': '批量替换文件内容',
            'status': '开发中'
        },
        {
            'icon': '🗑️',
            'name': '清理重复',
            'description': '查找并清理重复文件',
            'status': '开发中'
        }
    ]

    # 渲染工具卡片
    cols = st.columns(3)

    for i, tool in enumerate(tools):
        with cols[i % 3]:
            st.markdown(f"""
            <div class="tool-card">
                <div class="tool-icon">{tool['icon']}</div>
                <div class="tool-name">{tool['name']}</div>
                <div class="tool-description">{tool['description']}</div>
                <div class="tool-status">{tool['status']}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    st.info("💡 工具箱功能正在开发中,敬请期待!")
