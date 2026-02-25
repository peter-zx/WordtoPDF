# 云端部署指南

## 部署方案对比

### 方案1: 本地部署（当前方案）

**适用场景：**
- ✅ 本地使用
- ✅ 私有服务器
- ✅ 并发量小（<10用户）
- ✅ 成本敏感型

**技术栈：**
- LibreOffice
- Python多进程
- 无需额外服务

**优势：**
- ✅ 免费
- ✅ 稳定可靠
- ✅ 无需外网
- ✅ 数据安全

**劣势：**
- ❌ 不适合大规模云端部署
- ❌ 需要安装LibreOffice
- ❌ 资源占用较大

---

### 方案2: Docker部署（推荐用于私有云）

**适用场景：**
- ✅ 有Docker环境
- ✅ 私有云部署
- ✅ 中小规模应用
- ✅ 需要环境隔离

**Dockerfile示例：**
```dockerfile
FROM ubuntu:22.04

# 安装LibreOffice
RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# 安装Python
RUN apt-get install -y \
    python3 \
    python3-pip \
    --no-install-recommends

# 复制应用代码
COPY . /app
WORKDIR /app

# 安装依赖
RUN pip3 install -r requirements.txt

# 暴露端口
EXPOSE 8501

# 启动应用
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

**docker-compose.yml示例：**
```yaml
version: '3'
services:
  app:
    build: .
    ports:
      - "8501:8501"
    volumes:
      - ./data:/app/data
    environment:
      - TZ=Asia/Shanghai
    restart: unless-stopped
```

**部署命令：**
```bash
# 构建镜像
docker-compose build

# 启动服务
docker-compose up -d

# 查看日志
docker-compose logs -f
```

**优势：**
- ✅ 环境隔离
- ✅ 易于部署和扩展
- ✅ 支持水平扩展
- ✅ 数据安全

**劣势：**
- ❌ 每个容器资源占用大（~200MB）
- ❌ 不适合高并发
- ❌ 需要Docker环境

---

### 方案3: CloudConvert API（推荐用于公有云）⭐

**适用场景：**
- ✅ 大规模云端部署
- ✅ Serverless架构
- ✅ 高并发需求
- ✅ 对成本不敏感

**CloudConvert特点：**
- ✅ 专业的文档转换API
- ✅ 高可用（99.9% SLA）
- ✅ 支持所有格式
- ✅ 按使用量付费
- ✅ 免费额度：每月25次转换

**定价：**
- 免费版：25次/月
- 付费版：$0.01/次起（批量购买更便宜）

**集成示例：**

```python
from services.batch_pdf_service import BatchPDFService

# 使用CloudConvert API
results = BatchPDFService.batch_convert_with_structure(
    structure,
    output_dir,
    mode=BatchPDFService.MODE_CLOUDCONVERT,
    api_key="your-api-key"
)
```

**优势：**
- ✅ 无需安装任何软件
- ✅ 高可用，高并发
- ✅ 按需付费，成本可控
- ✅ 适合Serverless架构
- ✅ 支持大规模部署

**劣势：**
- ❌ 需要付费（但成本可控）
- ❌ 依赖第三方服务
- ❌ 数据需要上传到API

**成本估算：**
- 小规模应用（<100转换/月）：免费
- 中等规模（1000转换/月）：约$10
- 大规模（10000转换/月）：约$100

---

### 方案4: 混合方案（最灵活）

**适用场景：**
- ✅ 需要成本优化
- ✅ 有本地服务器
- ✅ 需要备份方案

**实现方式：**
```python
# 自动选择方案
results = BatchPDFService.batch_convert_with_structure(
    structure,
    output_dir,
    mode=BatchPDFService.MODE_AUTO,  # 自动选择
    api_key="your-api-key"  # CloudConvert备用
)
```

**工作流程：**
1. 优先尝试LibreOffice（免费）
2. LibreOffice失败时，自动切换到CloudConvert
3. 确保所有文件都能成功转换

**优势：**
- ✅ 成本最优
- ✅ 稳定性最高
- ✅ 灵活性最强

---

## 部署建议

### 阶段1: 本地开发/测试
**推荐：LibreOffice**
- 免费、快速、稳定
- 适合开发和测试

### 阶段2: 小规模生产（<100用户/天）
**推荐：Docker + LibreOffice**
- 易于部署
- 成本低
- 稳定可靠

### 阶段3: 大规模生产（>100用户/天）
**推荐：CloudConvert API**
- 高可用
- 高并发
- 按需付费

### 阶段4: 成本优化
**推荐：混合方案**
- LibreOffice为主
- CloudConvert为备
- 成本最优

---

## 服务器配置要求

### LibreOffice方案

**最低配置：**
- CPU: 2核
- 内存: 4GB
- 磁盘: 20GB

**推荐配置：**
- CPU: 4核
- 内存: 8GB
- 磁盘: 50GB

### CloudConvert方案

**最低配置：**
- CPU: 1核
- 内存: 1GB
- 磁盘: 10GB

**推荐配置：**
- CPU: 2核
- 内存: 2GB
- 磁盘: 20GB

---

## 安全建议

### LibreOffice方案
- ✅ 数据在本地处理，不上传
- ✅ 适合处理敏感文档
- ✅ 完全控制数据流

### CloudConvert方案
- ⚠️ 数据需要上传到API
- ⚠️ 需要遵守GDPR等隐私法规
- ⚠️ 建议对敏感数据进行加密

---

## 性能对比

| 方案 | 成本 | 并发 | 稳定性 | 适用场景 |
|------|------|------|--------|----------|
| LibreOffice | 免费 | 低 | 高 | 本地/私有服务器 |
| Docker | 低 | 中 | 高 | 私有云 |
| CloudConvert | 按需 | 高 | 很高 | 公有云/Serverless |
| 混合方案 | 低 | 中 | 很高 | 成本优化 |

---

## 总结

**对于云端部署，我的建议是：**

1. **小规模应用**：使用Docker + LibreOffice
2. **大规模应用**：使用CloudConvert API
3. **成本敏感**：使用混合方案

**当前代码已经支持所有方案，你只需要：**
1. 本地：直接使用（LibreOffice模式）
2. Docker：使用Dockerfile部署
3. 云端：配置CloudConvert API密钥

代码会自动选择最优方案！
