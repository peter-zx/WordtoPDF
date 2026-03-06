# 文档整理工具 - TypeScript版本

## 项目概述
基于Node.js和TypeScript的文档整理工具，提供Word转PDF、Excel文件夹结构创建、文件批量复制等功能。

## 技术栈
- **后端**: Node.js + TypeScript + Express
- **前端**: React + TypeScript + Ant Design
- **文件处理**: 
  - Word转PDF: LibreOffice/云服务API
  - Excel处理: exceljs
  - 文件操作: fs-extra
- **数据库**: SQLite (可选，用于记录操作历史)

## 项目结构
```
ts-project/
├── backend/                 # 后端服务
│   ├── src/
│   │   ├── controllers/     # 控制器
│   │   ├── services/        # 业务逻辑
│   │   ├── utils/           # 工具函数
│   │   └── app.ts          # 应用入口
│   ├── package.json
│   └── tsconfig.json
├── frontend/                # 前端界面
│   ├── src/
│   │   ├── components/     # 组件
│   │   ├── pages/          # 页面
│   │   └── App.tsx
│   ├── package.json
│   └── tsconfig.json
├── shared/                  # 共享类型定义
│   └── types.ts
└── package.json            # 根项目配置
```

## 核心功能模块

### 1. Word转PDF模块
- 支持.doc/.docx格式转换
- 批量转换，支持文件夹结构保持
- 转换进度实时显示

### 2. Excel处理模块
- 解析Excel表格结构
- 自动创建文件夹层级
- 支持复杂嵌套结构

### 3. 文件管理模块
- 文件扫描和筛选
- 批量复制操作
- 操作历史记录

## 部署方式
- 本地部署: Node.js环境
- 云服务器: Docker容器化部署
- 静态文件托管: 前后端分离部署

## 开发计划
1. ✅ 项目架构设计
2. 🔄 后端核心功能实现
3. ⏳ 前端界面开发
4. ⏳ 功能集成测试
5. ⏳ 部署配置