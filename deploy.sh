#!/bin/bash

# 运营成本管理看板部署脚本

echo "🚀 开始部署运营成本管理看板..."

# 检查 Docker 是否安装
if ! command -v docker &> /dev/null; then
    echo "❌ Docker 未安装，请先安装 Docker"
    exit 1
fi

# 检查 Docker Compose 是否安装
if ! command -v docker-compose &> /dev/null; then
    echo "❌ Docker Compose 未安装，请先安装 Docker Compose"
    exit 1
fi

# 创建日志目录
mkdir -p logs

# 停止现有容器
echo "🛑 停止现有容器..."
docker-compose down

# 构建并启动服务
echo "🔨 构建并启动服务..."
docker-compose up -d --build

# 等待服务启动
echo "⏳ 等待服务启动..."
sleep 10

# 检查服务状态
if docker-compose ps | grep -q "Up"; then
    echo "✅ 部署成功！"
    echo "📊 应用访问地址: http://localhost:8501"
    echo "📋 查看日志: docker-compose logs -f"
    echo "🛑 停止服务: docker-compose down"
else
    echo "❌ 部署失败，请检查日志: docker-compose logs"
    exit 1
fi