# WordtoPDF 项目文档

## 项目概述
Word转PDF批量转换工具，支持Windows本地环境和Linux服务器环境。

## 技术栈
- **前端框架**: Streamlit
- **后端**: Python
- **文档处理**: 
  - Windows: Microsoft Word (pywin32)
  - Linux: LibreOffice
- **依赖管理**: pip

## 项目结构

```
WordtoPDF/
├── app.py                    # Streamlit应用入口
├── requirements.txt          # Python依赖列表
├── pages/                    # Streamlit多页面
│   ├── 批量PDF转换.py       # 批量PDF转换页面（有问题）
│   ├── 文档格式转换.py       # 文档格式转换页面
│   ├── 文件夹生成.py         # 文件夹生成页面
│   ├── 工具箱.py             # 工具箱主页
│   └── __init__.py
├── components/               # 可复用组件
│   ├── folder_tree.py        # 文件夹树组件（Windows模式用）
│   ├── folder_selector.py    # 文件夹选择器
│   └── sidebar.py            # 侧边栏组件
├── services/                 # 业务逻辑服务
│   ├── batch_pdf_service.py  # 批量PDF转换服务
│   ├── word_converter.py     # Word转PDF转换器（跨平台）
│   ├── folder_service.py     # 文件夹服务
│   ├── file_service.py       # 文件服务
│   └── folder_scanner.py     # 文件夹扫描器
├── assets/                   # 静态资源
│   └── images/               # 图片资源
├── config.py                 # 配置文件
├── push.bat                  # 推送脚本
├── build.bat                 # 打包脚本
└── README.md                 # 项目说明
```

## 核心功能

### 1. 批量PDF转换 (`pages/批量PDF转换.py`)
**当前状态**: 有问题，需要修复

**功能**:
- 上传ZIP文件（服务器模式）
- 解压并扫描Word文件
- 选择要转换的文件
- 批量转换为PDF
- 下载转换结果

**问题**:
- 文件列表未按文件夹结构树展示
- Checkbox状态同步问题
- 重名文件处理缺失

**关键代码**:
```python
# 上传ZIP文件
uploaded_zip = st.file_uploader("选择ZIP文件", type=['zip'])

# 解压并获取文件列表
with zipfile.ZipFile(zip_path, 'r') as zipf:
    for member in zipf.namelist():
        # 处理文件名编码
        filename = member.encode('cp437').decode('gbk')

# Checkbox选择
new_value = st.checkbox(f"📄 {file_info['name']}", value=is_selected, key=f"file_{i}")
```

### 2. 文档格式转换 (`pages/文档格式转换.py`)
**状态**: 正常

**功能**: RTF、ODT、HTML、TXT等格式转换为DOCX

### 3. 文件夹生成 (`pages/文件夹生成.py`)
**状态**: 正常

**功能**: 根据列表批量生成文件夹并复制模板文件

### 4. 工具箱 (`pages/工具箱.py`)
**状态**: 正常

**功能**: 3行4列卡片布局，展示所有工具

## 服务层

### word_converter.py
**跨平台Word转PDF**

```python
def convert_single_file(input_path: str, output_path: str) -> Tuple[bool, str]:
    """转换单个文件为PDF - 自动选择最佳方法"""
    if is_windows():
        return convert_with_word(input_path, output_path)  # 使用MS Word
    else:
        return convert_with_libreoffice(input_path, output_dir)  # 使用LibreOffice
```

### batch_pdf_service.py
**批量PDF转换服务**

```python
def scan_folder_structure(folder_path: str) -> Dict:
    """扫描文件夹结构"""
    # 返回嵌套的文件夹结构
    return {
        "name": "文件夹名",
        "path": "路径",
        "type": "folder",
        "children": [...],  # 子文件夹
        "docx_files": [...]  # Word文件列表
    }
```

### folder_tree.py
**文件夹树组件（Windows模式）**

```python
class FolderTreeComponent:
    @staticmethod
    def render_folder_tree(structure: Dict):
        """渲染文件夹树"""
        # 使用expander实现展开/折叠
        # 使用checkbox实现选择功能
```

## Session State管理

### 批量PDF转换页面使用的状态
```python
st.session_state.uploaded_files = []      # 上传的文件列表
st.session_state.file_selection = {}      # 文件选择状态 {path: bool}
st.session_state.converted_files = []     # 转换结果
st.session_state.show_results = False     # 是否显示结果
```

## 跨平台支持

### Windows模式
- 使用tkinter选择文件夹
- 使用Microsoft Word转换
- 输出到本地目录

### Linux/服务器模式
- 上传ZIP文件
- 使用LibreOffice转换
- 下载转换结果（ZIP）

## 当前待修复问题

### 批量PDF转换页面 (`pages/批量PDF转换.py`)

1. **文件列表显示**
   - 需要显示文件夹结构树
   - 支持展开/折叠
   - 显示文件夹层级

2. **Checkbox状态管理**
   - 当前状态同步有问题
   - 勾选后计数不更新
   - 需要更稳定的状态管理方案

3. **重名文件处理**
   - 相同文件名导致覆盖
   - 需要保留文件夹结构
   - 或添加唯一后缀

4. **下载打包**
   - 保持文件夹结构
   - 确保所有文件都能保存

## 开发规范

### Git提交
- 使用 `push.bat` 脚本推送
- 不自动推送，等待指令
- 填写清晰的版本说明

### 测试
- 本地测试使用 `streamlit run app.py`
- 服务器测试需要推送到远程
- 注意Windows和Linux环境的差异

## 依赖说明

### requirements.txt
```
streamlit>=1.28.0
pandas>=2.0.0
openpyxl>=3.1.0
python-docx>=0.8.11
beautifulsoup4>=4.12.0
striprtf>=0.0.26
pywin32>=306          # Windows专用
comtypes>=1.2.0       # Windows专用
docx2pdf>=0.1.8      # Windows专用
```

## 服务器部署

### 服务信息
- IP: 122.51.231.239
- 端口: 8501
- 用户: ubuntu
- 服务: wordtopdf

### 部署方式
- 使用systemd服务
- 使用LibreOffice转换
- 持久化运行

## 关键文件说明

### app.py
Streamlit应用入口，定义多页面结构

### requirements.txt
Python依赖列表，区分Windows和Linux依赖

### push.bat
代码推送脚本，等待用户输入版本说明

### services/word_converter.py
跨平台Word转PDF转换器，自动选择转换方法

### components/folder_tree.py
文件夹树组件，Windows模式使用

## 注意事项

1. **编码问题**: ZIP文件名需要处理cp437到gbk编码
2. **状态管理**: Streamlit的session_state在rerun时会重置
3. **跨平台**: Windows和Linux的转换方式不同
4. **文件路径**: Windows使用反斜杠，Linux使用正斜杠
5. **临时文件**: 使用tempfile模块创建临时目录，用完清理

## 快速开始

### 本地运行
```bash
cd WordtoPDF
streamlit run app.py
```

### 服务器部署
```bash
# 连接服务器
ssh ubuntu@122.51.231.239

# 重启服务
sudo systemctl restart wordtopdf

# 查看日志
sudo journalctl -u wordtopdf -f
```

## 待优化功能

1. 批量PDF转换页面的文件夹树显示
2. Checkbox状态的稳定管理
3. 重名文件的处理
4. 下载时保持文件夹结构
5. 添加转换进度条（已部分实现）
6. 添加转换历史记录
