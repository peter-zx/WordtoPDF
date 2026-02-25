"""
文件夹树组件 - 严格按照架构实现
"""

import streamlit as st
from typing import Dict, List
import json
import os
import tempfile


class FolderTreeComponent:
    """文件夹树组件 - 严格架构实现"""

    _cache_file = None

    @staticmethod
    def _get_cache_file() -> str:
        if FolderTreeComponent._cache_file is None:
            FolderTreeComponent._cache_file = os.path.join(
                tempfile.gettempdir(), "pdf_converter_selected.json"
            )
        return FolderTreeComponent._cache_file

    @staticmethod
    def _load_selected() -> dict:
        """加载状态字典 - 返回 {path: bool}"""
        try:
            cache_file = FolderTreeComponent._get_cache_file()
            if os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data if isinstance(data, dict) else {}
        except:
            pass
        return {}

    @staticmethod
    def _save_selected(selected_dict: dict):
        """保存状态字典"""
        try:
            cache_file = FolderTreeComponent._get_cache_file()
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(selected_dict, f)
        except:
            pass

    @staticmethod
    def get_selected_files() -> List[str]:
        """获取选中文件列表"""
        selected_dict = FolderTreeComponent._load_selected()
        return [path for path, checked in selected_dict.items() if checked]

    @staticmethod
    def clear_cache():
        """清除所有选择"""
        FolderTreeComponent._save_selected({})

    @staticmethod
    def _get_all_file_paths(structure: Dict) -> List[str]:
        """获取所有文件路径"""
        paths = []
        for f in structure.get("docx_files", []):
            paths.append(f["path"])
        for child in structure.get("children", []):
            paths.extend(FolderTreeComponent._get_all_file_paths(child))
        return paths

    @staticmethod
    def render_selection_controls(structure: Dict):
        """渲染控制按钮"""
        all_paths = FolderTreeComponent._get_all_file_paths(structure)
        total = len(all_paths)
        
        selected_dict = FolderTreeComponent._load_selected()
        selected_count = sum(1 for p in all_paths if selected_dict.get(p, False))

        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])

        with col1:
            if st.button("全选", key="btn_all", use_container_width=True):
                # 全选：所有文件设为True
                new_dict = {p: True for p in all_paths}
                FolderTreeComponent._save_selected(new_dict)
                st.rerun()

        with col2:
            if st.button("反选", key="btn_invert", use_container_width=True):
                # 反选
                new_dict = {}
                for p in all_paths:
                    new_dict[p] = not selected_dict.get(p, False)
                FolderTreeComponent._save_selected(new_dict)
                st.rerun()

        with col3:
            if st.button("清空", key="btn_clear", use_container_width=True):
                # 清空：所有文件设为False
                new_dict = {p: False for p in all_paths}
                FolderTreeComponent._save_selected(new_dict)
                st.rerun()

        with col4:
            st.metric("已选择文件", f"{selected_count}/{total}")

    @staticmethod
    def render_folder_tree(structure: Dict) -> None:
        """渲染文件夹树"""
        # 注入样式
        st.markdown("""
        <style>
            .stCheckbox > label { font-size: 13px !important; }
            .stCheckbox { margin: 0 !important; padding: 0 !important; }
            div[data-testid="stExpander"] { margin: 2px 0 !important; }
            .streamlit-expanderHeader { font-size: 13px !important; padding: 4px !important; }
            .streamlit-expanderContent { padding: 0 0 0 16px !important; border-left: 2px solid #ddd !important; }
        </style>
        """, unsafe_allow_html=True)

        # 渲染根目录文件
        for f in structure.get("docx_files", []):
            FolderTreeComponent._render_file(f)

        # 渲染子文件夹
        for idx, child in enumerate(structure.get("children", [])):
            FolderTreeComponent._render_folder(child, 0, idx)

    @staticmethod
    def _render_file(file_info: Dict) -> None:
        """渲染单个文件 - 使用callback"""
        file_path = file_info["path"]
        file_name = file_info["name"]
        
        # 从缓存加载当前状态
        selected_dict = FolderTreeComponent._load_selected()
        is_checked = selected_dict.get(file_path, False)
        
        # 定义callback函数
        def on_file_change():
            # 获取当前状态
            current = FolderTreeComponent._load_selected()
            # 获取checkbox的新值
            new_value = st.session_state[file_key]
            # 更新字典
            current[file_path] = new_value
            # 保存到文件
            FolderTreeComponent._save_selected(current)
        
        file_key = f"file_cb_{hash(file_path) % 999999}"
        
        # 渲染checkbox，使用callback
        st.checkbox(
            f"📄 {file_name}", 
            value=is_checked, 
            key=file_key,
            on_change=on_file_change
        )

    @staticmethod
    def _render_folder(structure: Dict, level: int, idx: int) -> None:
        """渲染文件夹 - 使用callback"""
        folder_name = structure["name"]
        folder_path = structure["path"]

        # 获取该文件夹下所有文件路径
        all_paths = FolderTreeComponent._get_all_file_paths(structure)
        total = len(all_paths)
        
        # 从缓存加载状态
        selected_dict = FolderTreeComponent._load_selected()
        selected_count = sum(1 for p in all_paths if selected_dict.get(p, False))

        # 计算文件夹选中状态
        if selected_count == total and total > 0:
            status = "✅"
            folder_checked = True
        elif selected_count > 0:
            status = f"({selected_count}/{total})"
            folder_checked = False
        else:
            status = f"({total})"
            folder_checked = False

        indent = "　" * level
        folder_key = f"folder_cb_{level}_{idx}_{hash(folder_path) % 999999}"

        # 定义callback函数
        def on_folder_change():
            # 获取当前状态
            current = FolderTreeComponent._load_selected()
            # 获取checkbox的新值
            new_value = st.session_state[folder_key]
            # 更新所有子文件
            for p in all_paths:
                current[p] = new_value
            # 保存到文件
            FolderTreeComponent._save_selected(current)
        
        # 渲染文件夹checkbox，使用callback
        st.checkbox(
            f"{indent}📁 {folder_name} {status}",
            value=folder_checked,
            key=folder_key,
            on_change=on_folder_change
        )

        # 渲染内容
        has_content = structure.get("docx_files") or structure.get("children")
        if has_content:
            # 有选中内容时展开
            expanded = selected_count > 0

            with st.expander("▼", expanded=expanded):
                # 渲染文件
                for f in structure.get("docx_files", []):
                    FolderTreeComponent._render_file(f)

                # 递归渲染子文件夹
                for child_idx, child in enumerate(structure.get("children", [])):
                    FolderTreeComponent._render_folder(child, level + 1, child_idx)
