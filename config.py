"""
配置文件 - 统一管理配置和常量
"""

# 页面配置
PAGE_CONFIG = {
    "page_title": "文件批量处理工具箱",
    "page_icon": "🛠️",
    "layout": "wide",
    "initial_sidebar_state": "collapsed"
}

# 支持的文件格式
SUPPORTED_FORMATS = {
    "template": ["*"],  # 模板文件支持所有格式
    "text": ["txt"],
    "document": ["doc", "docx", "rtf", "odt", "pdf"],
    "convertible": ["rtf", "odt", "html", "htm"]
}

# 文件夹名称非法字符
INVALID_FOLDER_CHARS = '<>:"/\\|?*'

# UI配置
UI_CONFIG = {
    "colors": {
        "primary": "#0068C9",
        "success": "#00A651",
        "warning": "#FFB800",
        "error": "#D0021B"
    },
    "max_preview_lines": 10
}

# 工具模块配置
TOOLS_CONFIG = {
    "folder_generator": {
        "name": "文件夹批量生成",
        "icon": "📁",
        "description": "根据文本批量创建文件夹并复制模板文件"
    },
    "format_converter": {
        "name": "文档格式转换",
        "icon": "🔄",
        "description": "将RTF等格式转换为DOCX格式"
    },
    "batch_pdf": {
        "name": "批量DOCX转PDF",
        "icon": "📄",
        "description": "批量将DOCX文件转换为PDF格式"
    }
}
