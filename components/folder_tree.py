"""
文件夹树组件 - 使用session_state存储状态（支持多用户）
"""

import streamlit as st
from typing import Dict, List
import hashlib


class FolderTreeComponent:
    """文件夹树组件 - 使用session_state存储"""

    @staticmethod
    def _get_session_key() -> str:
        """获取当前会话的存储key"""
        return "folder_tree_selected"

    @staticmethod
    def _load_selected() -> dict:
        """加载状态字典 - 返回 {path: bool}"""
        key = FolderTreeComponent._get_session_key()
        if key not in st.session_state:
            st.session_state[key] = {}
        return st.session_state[key]

    @staticmethod
    def _save_selected(selected_dict: dict):
        """保存状态字典"""
        key = FolderTreeComponent._get_session_key()
        st.session_state[key] = selected_dict

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
                new_dict = dict(selected_dict)
                for p in all_paths:
                    new_dict[p] = True
                FolderTreeComponent._save_selected(new_dict)
                st.rerun()

        with col2:
            if st.button("反选", key="btn_invert", use_container_width=True):
                new_dict = dict(selected_dict)
                for p in all_paths:
                    new_dict[p] = not selected_dict.get(p, False)
                FolderTreeComponent._save_selected(new_dict)
                st.rerun()

        with col3:
            if st.button("清空", key="btn_clear", use_container_width=True):
                new_dict = dict(selected_dict)
                for p in all_paths:
                    new_dict[p] = False
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
    def _get_stable_key(path: str) -> str:
        """生成稳定的key"""
        # 使用路径的hash生成稳定的key
        return f"cb_{hashlib.md5(path.encode()).hexdigest()[:8]}"

    @staticmethod
    def _render_file(file_info: Dict) -> None:
        """渲染单个文件"""
        file_path = file_info["path"]
        file_name = file_info["name"]
        
        # 从session_state加载当前状态
        selected_dict = FolderTreeComponent._load_selected()
        is_checked = selected_dict.get(file_path, False)
        
        # 生成稳定的key
        file_key = FolderTreeComponent._get_stable_key(file_path)
        
        # 如果session_state中没有这个key，初始化它
        if file_key not in st.session_state:
            st.session_state[file_key] = is_checked
        
        # 定义callback函数
        def on_file_change():
            current = FolderTreeComponent._load_selected()
            new_value = st.session_state[file_key]
            current[file_path] = new_value
            FolderTreeComponent._save_selected(current)
        
        # 渲染checkbox
        st.checkbox(
            f"📄 {file_name}", 
            value=is_checked, 
            key=file_key,
            on_change=on_file_change
        )

    @staticmethod
    def _render_folder(structure: Dict, level: int, idx: int) -> None:
        """渲染文件夹"""
        folder_name = structure["name"]
        folder_path = structure["path"]

        # 获取该文件夹下所有文件路径
        all_paths = FolderTreeComponent._get_all_file_paths(structure)
        total = len(all_paths)
        
        # 从session_state加载状态
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
        folder_key = f"folder_{level}_{idx}_{hashlib.md5(folder_path.encode()).hexdigest()[:8]}"
        
        # 如果session_state中没有这个key，初始化它
        if folder_key not in st.session_state:
            st.session_state[folder_key] = folder_checked

        # 定义callback函数
        def on_folder_change():
            current = FolderTreeComponent._load_selected()
            new_value = st.session_state[folder_key]
            for p in all_paths:
                current[p] = new_value
            FolderTreeComponent._save_selected(current)
        
        # 渲染文件夹checkbox
        st.checkbox(
            f"{indent}📁 {folder_name} {status}",
            value=folder_checked,
            key=folder_key,
            on_change=on_folder_change
        )

        # 渲染内容
        has_content = structure.get("docx_files") or structure.get("children")
        if has_content:
            expanded = selected_count > 0

            with st.expander("▼", expanded=expanded):
                # 渲染文件
                for f in structure.get("docx_files", []):
                    FolderTreeComponent._render_file(f)

                # 递归渲染子文件夹
                for child_idx, child in enumerate(structure.get("children", [])):
                    FolderTreeComponent._render_folder(child, level + 1, child_idx)
