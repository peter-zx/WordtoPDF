# 文件批量处理工具箱

一个基于Python Streamlit的专业级离线Web工具箱,提供多种文件批量处理功能。

## 功能特性

### 🛠️ 核心工具

1. **📁 文件夹批量生成**
   - 根据文本批量创建文件夹
   - 支持从TXT文件读取或直接粘贴
   - 自动复制模板文件到每个文件夹

2. **🔄 文档格式转换**
   - RTF → DOCX
   - ODT → DOCX
   - HTML → DOCX
   - 旧版DOC → DOCX
   - 支持批量转换

3. **📄 批量PDF转换**
   - DOCX → PDF
   - 支持多级文件夹递归
   - 保持目录结构
   - 批量处理

### ✨ 界面特点

- 🎨 三栏布局,一目了然
- 📱 响应式设计
- 🎯 顶部常用按钮快捷切换
- 📊 实时进度和结果反馈
- 🔒 纯离线运行,数据安全
- 🏗️ 模块化架构,易于扩展

## 快速开始

### Windows用户

#### 方式1: 自动安装和运行

```bash
# 1. 双击 install.bat 进行首次安装
# 2. 安装完成后双击 run.bat 运行程序
```

#### 方式2: 手动安装

```bash
# 1. 创建虚拟环境
python -m venv venv

# 2. 激活虚拟环境
venv\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
streamlit run app.py
```

### Linux/Mac用户

```bash
# 1. 创建虚拟环境
python3 -m venv venv

# 2. 激活虚拟环境
source venv/bin/activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
streamlit run app.py

# 或使用启动脚本
chmod +x setup.sh
./setup.sh
```

## 使用指南

### 文件夹批量生成工具

1. **步骤1: 输入文件夹名称**
   - 上传TXT文件或直接粘贴文本
   - 每行一个文件夹名称

2. **步骤2: 选择目标文件夹和模板**
   - 指定创建文件夹的位置
   - 选择要复制的模板文件(支持多选)

3. **步骤3: 执行批量创建**
   - 点击"开始批量创建文件夹"
   - 查看操作结果

### 文档格式转换工具

1. **步骤1: 选择输入文件**
   - 支持RTF、ODT、HTML、DOC格式
   - 可批量选择多个文件

2. **步骤2: 选择输出目录**
   - 指定转换后DOCX文件的保存位置

3. **步骤3: 执行转换**
   - 点击"开始转换"
   - 查看转换结果

### 批量PDF转换工具

1. **步骤1: 选择输入**
   - 选择DOCX文件或包含DOCX的文件夹
   - 支持递归查找子文件夹

2. **步骤2: 设置输出**
   - 指定PDF保存目录
   - 选择是否保持目录结构

3. **步骤3: 执行转换**
   - 点击"开始转换为PDF"
   - 查看转换结果

## 项目结构

```
wenjianjia/
├── venv/                      # 虚拟环境
├── app.py                     # 主程序入口
├── config.py                  # 配置文件
├── utils.py                   # 工具函数模块
├── ui_components.py           # UI组件模块
├── folder_generator_module.py # 文件夹生成工具
├── format_converter.py        # 格式转换核心
├── format_converter_module.py # 格式转换界面
├── pdf_converter.py           # PDF转换核心
├── pdf_converter_module.py    # PDF转换界面
├── requirements.txt           # 依赖列表
├── install.bat                # Windows安装脚本
├── run.bat                    # Windows启动脚本
├── setup.sh                   # Linux/Mac启动脚本
├── example.txt                # 示例文件
└── README.md                  # 说明文档
```

## 技术栈

- **Python 3.7+**
- **Streamlit** - WebUI框架
- **pypandoc** - 文档格式转换
- **docx2pdf** - DOCX转PDF
- **comtypes** - Windows PDF转换备用方案
- **venv** - 虚拟环境隔离

## 系统要求

### 基础要求
- Windows / macOS / Linux
- Python 3.7+
- 现代浏览器(Chrome, Firefox, Edge等)

### 格式转换额外要求

**文档格式转换(RTF/ODT/HTML):**
- 需要安装pandoc: https://pandoc.org/installing.html
- 或使用Python包: `pip install pypandoc`

**PDF转换(DOCX → PDF):**

Windows:
```bash
pip install docx2pdf  # 推荐
# 或
pip install comtypes  # 备用方案
```

Linux/Mac:
```bash
# 需要先安装LibreOffice
sudo apt-get install libreoffice  # Ubuntu/Debian
brew install --cask libreoffice  # macOS

pip install docx2pdf
```

## 配置说明

### config.py
- 页面配置
- 支持的文件格式
- UI颜色主题
- 工具模块配置

### 自定义配置
可以修改`config.py`文件来自定义:
- 页面标题和图标
- 支持的文件格式
- UI颜色方案
- 工具模块名称和描述

## 故障排除

### 虚拟环境激活失败
```bash
# 确保Python已正确安装
python --version

# 删除并重新创建虚拟环境
rm -rf venv
python -m venv venv
```

### 依赖安装失败
```bash
# 使用国内镜像
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 单独安装失败的包
pip install pypandoc
pip install docx2pdf
```

### 文档格式转换失败
```bash
# 检查pandoc是否安装
pandoc --version

# 如果未安装,访问: https://pandoc.org/installing.html
```

### PDF转换失败
```bash
# Windows用户
pip install pywin32  # 可能需要

# Linux/Mac用户
# 确保LibreOffice已安装
libreoffice --version
```

### 端口被占用
```bash
# 指定其他端口启动
streamlit run app.py --server.port 8502
```

## 开发说明

### 添加新工具

1. 在`config.py`中添加工具配置
2. 创建工具模块文件(如`new_tool.py`)
3. 创建工具界面模块(如`new_tool_module.py`)
4. 在`app.py`中注册工具

### 代码结构

- **config.py**: 集中管理配置
- **utils.py**: 通用工具函数
- **ui_components.py**: 可复用UI组件
- ***_module.py**: 各工具界面逻辑
- **app.py**: 主程序入口和路由

### 模块化设计

所有功能模块独立封装,便于:
- 单独测试
- 代码复用
- 功能扩展
- 维护升级

## 注意事项

- 使用虚拟环境隔离依赖
- 所有操作在本地执行,数据不上传
- 目标文件夹必须存在且有写入权限
- 转换大文件时可能需要较长时间
- 建议先在小批量文件上测试

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request!
