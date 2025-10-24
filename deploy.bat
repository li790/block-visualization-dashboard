@echo off
chcp 65001 >nul

echo 🚀 开始部署运营成本管理看板...

REM 检查 Docker 是否安装
docker --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker 未安装，请先安装 Docker
    pause
    exit /b 1
)

REM 检查 Docker Compose 是否安装
docker-compose --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker Compose 未安装，请先安装 Docker Compose
    pause
    exit /b 1
)

REM 创建日志目录
if not exist logs mkdir logs

REM 停止现有容器
echo 🛑 停止现有容器...
docker-compose down

REM 构建并启动服务
echo 🔨 构建并启动服务...
docker-compose up -d --build

REM 等待服务启动
echo ⏳ 等待服务启动...
timeout /t 10 /nobreak >nul

REM 检查服务状态
docker-compose ps | findstr "Up" >nul
if errorlevel 1 (
    echo ❌ 部署失败，请检查日志: docker-compose logs
    pause
    exit /b 1
) else (
    echo ✅ 部署成功！
    echo 📊 应用访问地址: http://localhost:8501
    echo 📋 查看日志: docker-compose logs -f
    echo 🛑 停止服务: docker-compose down
)

pause