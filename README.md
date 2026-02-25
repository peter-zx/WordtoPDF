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

### Windows用户

#### 方式1: 完整自动安装（推荐）

```bash
# 1. 双击 install_full.bat 进行完整安装（包括系统依赖）
# 2. 安装完成后双击 run.bat 运行程序
```

#### 方式2: 基础安装

```bash
# 1. 双击 install.bat 进行基础安装
# 2. 手动安装 Pandoc（用于文档格式转换）
#    下载地址: https://pandoc.org/installing.html
#    或使用 winget install --id JohnMacFarlane.Pandoc
# 3. 安装完成后双击 run.bat 运行程序
```

#### 方式3: 手动安装

```bash
# 1. 创建虚拟环境
python -m venv venv

# 2. 激活虚拟环境
venv\Scripts\activate

# 3. 安装Python依赖
pip install -r requirements.txt

# 4. 安装 Pandoc（必需，用于文档格式转换）
#    方式1: winget install --id JohnMacFarlane.Pandoc
#    方式2: 从 https://pandoc.org/installing.html 下载安装

# 5. 运行程序
streamlit run app.py
```
