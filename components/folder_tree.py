"""
文件夹树组件 - 可视化选择文件夹
"""

import streamlit as st
from typing import Dict, List, Set


class FolderTreeComponent:
    """文件夹树组件"""

    def __init__(self):
        """初始化组件"""
        pass

    @staticmethod
    def render_folder_tree(
        structure: Dict,
        selected_files: Set[str],
        level: int = 0
    ) -> Set[str]:
        """
        渲染文件夹树，返回选中的文件列表

        Args:
            structure: 文件夹结构
            selected_files: 已选中的文件集合
            level: 缩进级别

        Returns:
            选中的文件集合
        """
        indent = "　" * level

        # 渲染当前文件夹
        folder_name = structure["name"]

        # 检查当前文件夹下的所有文件
        all_files = FolderTreeComponent._get_all_files_in_folder(structure)

        # 检查当前文件夹是否全部选中
        all_selected = all(
            file_info["path"] in selected_files
            for file_info in all_files
        )

        # 检查当前文件夹是否部分选中
        some_selected = any(
            file_info["path"] in selected_files
            for file_info in all_files
        )

        # 文件夹复选框
        if all_selected:
            folder_key = f"folder_{structure['path']}"
            folder_checked = st.checkbox(
                f"📁 {indent}{folder_name}",
                value=True,
                key=folder_key,
                help=f"包含 {len(all_files)} 个文件"
            )

            if not folder_checked:
                # 取消选中，移除所有文件
                for file_info in all_files:
                    selected_files.discard(file_info["path"])

        elif some_selected:
            # 部分选中，显示半选中状态（通过文本标记）
            st.markdown(f"📁 {indent}{folder_name} ⚠️ (部分选中: {sum(1 for f in all_files if f['path'] in selected_files)}/{len(all_files)})")

            # 提供全选/取消全选按钮
            col1, col2 = st.columns(2)
            with col1:
                if st.button(f"全选", key=f"select_all_{structure['path']}"):
                    for file_info in all_files:
                        selected_files.add(file_info["path"])
                    st.rerun()

            with col2:
                if st.button(f"取消", key=f"clear_{structure['path']}"):
                    for file_info in all_files:
                        selected_files.discard(file_info["path"])
                    st.rerun()

        else:
            # 未选中
            col1, col2 = st.columns([4, 1])
            with col1:
                st.markdown(f"📁 {indent}{folder_name}")
            with col2:
                if st.button("全选", key=f"select_all_{structure['path']}", use_container_width=True):
                    for file_info in all_files:
                        selected_files.add(file_info["path"])
                    st.rerun()

        # 渲染当前文件夹的文件
        if structure.get("docx_files"):
            for file_info in structure["docx_files"]:
                file_key = f"file_{file_info['path']}"
                is_selected = file_info["path"] in selected_files

                if is_selected:
                    checked = st.checkbox(
                        f"📄 {indent}　{file_info['name']}",
                        value=True,
                        key=file_key
                    )
                    if not checked:
                        selected_files.discard(file_info["path"])
                else:
                    checked = st.checkbox(
                        f"📄 {indent}　{file_info['name']}",
                        value=False,
                        key=file_key
                    )
                    if checked:
                        selected_files.add(file_info["path"])

        # 递归渲染子文件夹
        for child in structure.get("children", []):
            selected_files = FolderTreeComponent.render_folder_tree(
                child,
                selected_files,
                level + 1
            )

        return selected_files

    @staticmethod
    def _get_all_files_in_folder(structure: Dict) -> List[Dict]:
        """获取文件夹及其子文件夹中的所有文件"""
        files = []

        # 添加当前文件夹的文件
        files.extend(structure.get("docx_files", []))

        # 递归添加子文件夹的文件
        for child in structure.get("children", []):
            files.extend(FolderTreeComponent._get_all_files_in_folder(child))

        return files

    @staticmethod
    def render_selection_controls(
        structure: Dict,
        selected_files: Set[str]
    ):
        """
        渲染选择控制按钮（全选/反选/清空）

        Args:
            structure: 文件夹结构
            selected_files: 已选中的文件集合
        """
        all_files = FolderTreeComponent._get_all_files_in_folder(structure)

        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])

        with col1:
            if st.button("全选", key="select_all_global"):
                for file_info in all_files:
                    selected_files.add(file_info["path"])
                st.rerun()

        with col2:
            if st.button("反选", key="invert_selection"):
                for file_info in all_files:
                    if file_info["path"] in selected_files:
                        selected_files.discard(file_info["path"])
                    else:
                        selected_files.add(file_info["path"])
                st.rerun()

        with col3:
            if st.button("清空", key="clear_all"):
                selected_files.clear()
                st.rerun()

        with col4:
            selected_count = len(selected_files)
            total_count = len(all_files)
            st.metric(
                "已选择",
                f"{selected_count}/{total_count}",
                delta=f"{selected_count} 文件"
            )
