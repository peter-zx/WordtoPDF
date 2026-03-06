# -*- coding: utf-8 -*-
"""日志模块 - 支持错误日志保存"""

import os
import logging
import logging.handlers
from datetime import datetime


class Logger:
    """日志管理器"""
    
    def __init__(self, log_dir="logs", app_name="文档整理工具"):
        self.log_dir = log_dir
        self.app_name = app_name
        self._setup_logging()
    
    def _setup_logging(self):
        """设置日志配置"""
        # 创建日志目录
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # 日志文件名格式：年-月-日_app.log
        log_filename = f"{datetime.now().strftime('%Y-%m-%d')}_{self.app_name.replace(' ', '_')}.log"
        log_filepath = os.path.join(self.log_dir, log_filename)
        
        # 配置根日志记录器
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        
        # 清除现有处理器
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # 文件处理器 - 记录所有级别的日志
        file_handler = logging.handlers.RotatingFileHandler(
            log_filepath,
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        
        # 控制台处理器 - 只记录INFO及以上级别
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 日志格式
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # 添加处理器
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        # 记录启动信息
        self.info(f"{self.app_name} 启动")
        self.info(f"日志文件: {log_filepath}")
    
    def debug(self, message):
        """调试级别日志"""
        logging.debug(message)
    
    def info(self, message):
        """信息级别日志"""
        logging.info(message)
    
    def warning(self, message):
        """警告级别日志"""
        logging.warning(message)
    
    def error(self, message, exc_info=False):
        """错误级别日志"""
        logging.error(message, exc_info=exc_info)
    
    def critical(self, message):
        """严重错误级别日志"""
        logging.critical(message)
    
    def log_operation(self, operation, details):
        """记录操作日志"""
        self.info(f"操作: {operation} - {details}")
    
    def log_error_with_traceback(self, operation, error):
        """记录带堆栈跟踪的错误"""
        import traceback
        error_details = f"操作: {operation}, 错误: {str(error)}\n堆栈跟踪:\n{traceback.format_exc()}"
        self.error(error_details)
    
    def get_recent_logs(self, lines=50):
        """获取最近的日志内容"""
        try:
            log_filename = f"{datetime.now().strftime('%Y-%m-%d')}_{self.app_name.replace(' ', '_')}.log"
            log_filepath = os.path.join(self.log_dir, log_filename)
            
            if os.path.exists(log_filepath):
                with open(log_filepath, 'r', encoding='utf-8') as f:
                    log_lines = f.readlines()
                return ''.join(log_lines[-lines:])
            else:
                return "暂无日志文件"
        except Exception as e:
            return f"读取日志失败: {str(e)}"


# 全局日志实例
logger = Logger()