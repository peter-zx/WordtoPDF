# -*- coding: utf-8 -*-
"""文档整理工具 - 主入口"""
import tkinter as tk
import traceback
from ui.main_window import MainWindow
from core.excel_parser import ExcelParser
from core.folder_manager import FolderManager
from core.file_copier import FileCopier
from core.word_converter import WordConverter
from core.logger import logger


def main():
    """主函数"""
    try:
        logger.info("文档整理工具启动")
        
        root = tk.Tk()

        # 创建核心组件
        excel_parser = ExcelParser()
        folder_manager = FolderManager()
        file_copier = FileCopier()
        word_converter = WordConverter()

        # 创建主窗口
        app = MainWindow(root)

        # 注入组件
        app.set_components(excel_parser, folder_manager, file_copier, word_converter)

        # 设置UI
        app.setup_ui()

        logger.info("UI初始化完成，开始运行")
        
        # 运行
        root.mainloop()
        
        logger.info("文档整理工具正常退出")
        
    except Exception as e:
        error_msg = f"程序启动失败: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        
        # 显示错误对话框
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        from tkinter import messagebox
        messagebox.showerror("启动错误", f"程序启动失败:\n{str(e)}\n\n请检查日志文件获取详细信息。")
        root.destroy()
        
        raise


if __name__ == "__main__":
    main()
