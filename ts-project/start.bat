@echo off
echo ====================================
echo   文档整理工具 - TypeScript版本
echo ====================================
echo.

echo [1/3] 安装后端依赖...
cd backend
if not exist node_modules (
    npm install
) else (
    echo 后端依赖已存在，跳过安装
)

echo [2/3] 安装前端依赖...
cd ..\frontend
if not exist node_modules (
    npm install
) else (
    echo 前端依赖已存在，跳过安装
)

echo [3/3] 启动开发服务器...
echo.
echo 请按以下步骤启动服务：
echo 1. 新开一个命令行窗口，执行: cd backend && npm run dev
echo 2. 新开一个命令行窗口，执行: cd frontend && npm run dev
echo.
echo 后端服务将运行在: http://localhost:3001
echo 前端服务将运行在: http://localhost:3000
echo.
pause