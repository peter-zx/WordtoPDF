# WordtoPDF 应用打包说明

## 打包步骤

### 1. 准备环境
确保已安装Python 3.8或更高版本

### 2. 安装依赖
```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 3. 执行打包
双击运行 `build.bat` 脚本，或在命令行执行：
```bash
build.bat
```

### 4. 获取安装包
打包完成后，可执行文件位于：
```
dist/WordtoPDF.exe
```

## 用户使用说明

### 发送给用户
将 `dist/WordtoPDF.exe` 文件发送给其他用户即可

### 运行应用
1. 双击 `WordtoPDF.exe` 启动应用
2. 浏览器会自动打开 http://127.0.0.1:8501
3. 开始使用应用

### 系统要求
- Windows 10/11 64位
- 需要安装 Microsoft Word (用于Word转PDF功能)

## 注意事项

1. **首次启动可能较慢**：因为需要解压内置的Python环境
2. **杀毒软件可能误报**：首次运行可能需要添加信任
3. **防火墙提示**：允许应用访问本地网络
4. **Word要求**：批量PDF转换功能需要安装Microsoft Word

## 打包选项

如需自定义打包，可编辑 `build.bat` 文件中的参数：

- `--onefile`: 打包成单个exe文件
- `--windowed`: 不显示控制台窗口
- `--icon`: 设置应用图标
- `--add-data`: 添加额外的文件/文件夹

## 故障排查

### 打包失败
- 检查Python版本是否 >= 3.8
- 确保所有依赖已正确安装
- 查看错误日志

### 运行报错
- 检查是否安装了Microsoft Word
- 查看临时文件夹权限
- 检查杀毒软件是否拦截
