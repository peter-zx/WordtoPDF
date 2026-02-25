# 技术栈总结

## 项目概述
文件批量处理工具箱 - 基于Python Streamlit的专业级离线Web工具箱

## 核心技术栈

### 1. 前端框架
- **Streamlit** (v1.54.0)
  - Python Web应用框架
  - 用于构建交互式用户界面
  - 支持实时数据更新和交互式组件

### 2. 后端语言
- **Python** (v3.10.9)
  - 主要开发语言
  - 用于业务逻辑处理和文件操作

### 3. 文档处理库

#### 3.1 Word文档处理
- **python-docx** (v1.2.0)
  - 创建和操作Word文档
  - 用于TXT/HTML转DOCX

- **pywin32** (v311)
  - Windows COM接口
  - 用于调用Microsoft Word/WPS Office
  - 实现RTF转DOCX、DOCX转PDF

#### 3.2 文档格式转换
- **pypandoc** (v1.16.2)
  - 通用文档转换工具
  - 支持多种文档格式互转
  - 依赖系统级Pandoc (v3.9)

- **striprtf** (v0.0.29)
  - RTF文本提取工具
  - 备用的RTF处理方案

#### 3.3 PDF转换
- **docx2pdf** (v0.1.8)
  - DOCX转PDF工具
  - 批量PDF转换核心依赖

- **comtypes** (v1.4.15)
  - Windows COM类型库
  - PDF转换的备用方案

### 4. 数据处理
- **pandas** (v2.3.3)
  - 数据处理库
  - 用于文件列表管理和数据操作

- **openpyxl** (v3.1.5)
  - Excel文件处理
  - 用于读取和写入Excel文件

### 5. HTML解析
- **beautifulsoup4** (v4.14.3)
  - HTML/XML解析库
  - 用于HTML转DOCX

### 6. 系统级依赖
- **Pandoc** (v3.9)
  - 系统级文档转换工具
  - 通过winget自动安装
  - 支持RTF、ODT、HTML等格式转换

### 7. 虚拟环境管理
- **venv**
  - Python虚拟环境
  - 环境隔离和依赖管理

## 项目架构

### 目录结构
```
WordtoPDF/
├── app.py                          # 主程序入口
├── config.py                       # 配置文件
├── requirements.txt                # Python依赖列表
├── install.bat                     # 基础安装脚本
├── install_full.bat                # 完整安装脚本（含系统依赖）
├── run.bat                         # Windows运行脚本
├── setup.sh                        # Linux/Mac运行脚本
├── README.md                       # 项目说明文档
├── TECH_STACK.md                   # 技术栈文档（本文件）
├── pages/                          # 页面模块
│   ├── __init__.py
│   ├── 文件夹生成.py              # 文件夹批量生成页面
│   ├── 工具箱.py                  # 工具箱首页
│   ├── 批量PDF转换.py            # 批量PDF转换页面
│   └── 文档格式转换.py            # 文档格式转换页面
├── services/                       # 业务逻辑服务
│   ├── __init__.py
│   ├── batch_pdf_service.py       # PDF批量转换服务
│   ├── file_service.py            # 文件操作服务
│   ├── folder_service.py          # 文件夹操作服务
│   └── format_converter_service.py # 文档格式转换服务
├── components/                     # UI组件
│   ├── __init__.py
│   ├── folder_selector.py         # 文件夹选择组件
│   └── sidebar.py                 # 侧边栏组件
├── assets/                        # 静态资源
└── venv/                          # Python虚拟环境
```

## 核心功能实现

### 1. 文件夹批量生成
- **技术**: 文件系统操作
- **实现**: 批量创建文件夹、复制模板文件
- **页面**: `pages/文件夹生成.py`
- **服务**: `services/folder_service.py`

### 2. 文档格式转换
- **技术**: COM接口、Pandoc、BeautifulSoup
- **支持格式**: RTF、ODT、HTML、TXT → DOCX
- **页面**: `pages/文档格式转换.py`
- **服务**: `services/format_converter_service.py`
- **核心方法**:
  - RTF转DOCX: 使用Word/WPS COM接口，指定GBK编码
  - ODT转DOCX: 使用pypandoc
  - HTML转DOCX: 使用BeautifulSoup解析
  - TXT转DOCX: 使用python-docx

### 3. 批量PDF转换
- **技术**: Word/WPS COM接口
- **功能**: DOCX → PDF，保持文件夹结构
- **页面**: `pages/批量PDF转换.py`
- **服务**: `services/batch_pdf_service.py`
- **特点**:
  - 支持多层文件夹递归
  - 保持原始目录结构
  - 实时进度显示
  - 使用tkinter文件夹选择对话框

## 部署和安装

### 系统要求
- **操作系统**: Windows / macOS / Linux
- **Python**: 3.7+
- **浏览器**: 现代浏览器（Chrome、Firefox、Edge等）

### 安装方式

#### 方式1: 完整自动安装（推荐）
```bash
# Windows
install_full.bat
```

#### 方式2: 手动安装
```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# 安装Python依赖
pip install -r requirements.txt

# 安装Pandoc（必需）
winget install --id JohnMacFarlane.Pandoc  # Windows
# 或访问 https://pandoc.org/installing.html 下载安装
```

### 运行方式
```bash
# Windows
run.bat

# Linux/Mac
./setup.sh

# 或直接运行
streamlit run app.py
```

## 技术亮点

### 1. 模块化设计
- 功能模块独立封装
- 易于维护和扩展
- 代码复用性高

### 2. 环境隔离
- 使用venv虚拟环境
- 避免依赖冲突
- 便于部署和迁移

### 3. 离线运行
- 所有操作在本地执行
- 数据不上传服务器
- 保证数据安全

### 4. 用户友好
- 标准的文件夹选择对话框
- 实时进度反馈
- 详细的错误提示

### 5. 跨平台支持
- Windows、macOS、Linux
- 自动检测系统环境
- 适配不同Office套件

## 性能优化

### 1. 单例模式
- Word/WPS COM对象使用单例模式
- 避免重复创建进程
- 提高转换效率

### 2. 批量处理
- 支持批量文件转换
- 减少重复操作
- 提升整体效率

### 3. 错误处理
- 完善的异常捕获
- 优雅的错误提示
- 避免程序崩溃

## 安全性

### 1. 数据本地化
- 所有文件操作在本地进行
- 不涉及网络传输
- 保护用户隐私

### 2. 输入验证
- 文件路径验证
- 格式检查
- 防止路径注入

### 3. 错误隔离
- 单个文件失败不影响其他文件
- 详细的错误日志
- 便于问题排查

## 未来扩展

### 可能的改进方向
1. 支持更多文档格式（PDF、EPUB等）
2. 添加文档合并功能
3. 支持云端存储集成
4. 添加批量重命名功能
5. 支持自定义转换规则

## 依赖版本信息

```
streamlit==1.54.0
pandas==2.3.3
openpyxl==3.1.5
python-docx==1.2.0
beautifulsoup4==4.14.3
pypandoc==1.16.2
docx2pdf==0.1.8
comtypes==1.4.15
pywin32==311
striprtf==0.0.29
```

## 系统依赖

```
Pandoc==3.9
Microsoft Word 或 WPS Office
```

## 许可证

MIT License

## 作者

CodeArts Agent

## 更新日志

### 2025-02-25
- 优化批量PDF转换页面，使用标准文件夹选择对话框
- 修复RTF中文编码问题，指定GBK编码参数
- 添加完整安装脚本，自动安装系统依赖
- 清理冗余代码，优化项目结构
- 更新技术栈文档
