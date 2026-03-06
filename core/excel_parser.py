# -*- coding: utf-8 -*-
"""Excel解析模块"""
import pandas as pd
import os
from .logger import logger


class ExcelParser:
    """Excel文件解析器"""

    def __init__(self):
        self.structure = {}

    def load_from_file(self, excel_path):
        """从Excel文件加载文件夹结构"""
        logger.info(f"开始加载Excel文件: {excel_path}")
        
        if not excel_path or not os.path.exists(excel_path):
            error_msg = f"Excel文件不存在: {excel_path}"
            logger.error(error_msg)
            raise FileNotFoundError(error_msg)

        try:
            df = pd.read_excel(excel_path)
            cols = df.columns.tolist()
            logger.info(f"Excel列名: {cols}")

            structure = {}
            row_count = 0
            for _, row in df.iterrows():
                col1 = str(row[cols[0]]).strip() if pd.notna(row[cols[0]]) else ""
                col2 = str(row[cols[1]]).strip() if pd.notna(row[cols[1]]) else ""
                col3 = str(row[cols[2]]).strip() if len(cols) > 2 and pd.notna(row[cols[2]]) else ""

                if not col1:
                    continue

                if col1 not in structure:
                    structure[col1] = {}

                if col2:
                    if col2 not in structure[col1]:
                        structure[col1][col2] = []
                    if col3:
                        structure[col1][col2].append(col3)
                        row_count += 1
                        
            logger.info(f"Excel解析完成，处理了 {row_count} 行数据")
            logger.info(f"生成结构: {list(structure.keys())}")
            
            self.structure = structure
            return structure
            
        except Exception as e:
            error_msg = f"Excel文件解析失败: {str(e)}"
            logger.error(error_msg)
            raise

    def get_structure(self):
        """获取文件夹结构"""
        return self.structure

    def clear(self):
        """清空结构"""
        self.structure = {}
