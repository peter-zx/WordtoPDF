"""
侧边栏组件
"""

import streamlit as st
from typing import List, Dict, Any


class Sidebar:
    """侧边栏组件"""

    @staticmethod
    def render(nav_items: List[Dict[str, Any]], current_page: str):
        """
        渲染侧边栏

        Args:
            nav_items: 导航项列表
                [
                    {
                        'key': '唯一标识',
                        'icon': '图标',
                        'label': '文字'
                    }
                ]
            current_page: 当前页面key
        """
        st.markdown('<div class="sidebar-nav">', unsafe_allow_html=True)

        for item in nav_items:
            is_active = item['key'] == current_page
            active_class = 'active' if is_active else ''

            # 使用按钮实现导航
            if st.button(
                f"{item['icon']} {item['label']}",
                key=f"nav_{item['key']}",
                use_container_width=True,
                type="primary" if is_active else "secondary"
            ):
                st.session_state.current_page = item['key']
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
