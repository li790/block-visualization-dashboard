# 运营成本管理看板部署指南

## 部署方案概述

本项目提供两种部署方案：
1. **Docker 部署**（推荐）- 适合服务器和生产环境
2. **直接部署** - 适合内网和小规模使用

## 方案一：Docker 部署（推荐）

### 环境要求
- Docker 20.10+
- Docker Compose 2.0+
- 2GB+ 可用内存
- 1GB+ 可用磁盘空间

### 快速部署

#### Windows 系统
```bash
# 1. 双击运行部署脚本
deploy.bat

# 2. 或手动执行命令
docker-compose up -d --build
```

#### Linux/Mac 系统
```bash
# 1. 给脚本执行权限
chmod +x deploy.sh

# 2. 运行部署脚本
./deploy.sh

# 3. 或手动执行命令
docker-compose up -d --build
```

### 访问应用
- **本地访问**: http://localhost:8501
- **局域网访问**: http://[服务器IP]:8501

### 常用命令

```bash
# 查看服务状态
docker-compose ps

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down

# 重启服务
docker-compose restart

# 更新服务
docker-compose pull && docker-compose up -d --build
```

## 方案二：直接部署

### 环境要求
- Python 3.7+
- pip 包管理器

### 部署步骤

#### 1. 安装 Python 依赖
```bash
pip install -r requirements.txt
```

#### 2. 启动应用
```bash
# 本地访问
streamlit run main.py

# 局域网访问
streamlit run main.py --server.address 0.0.0.0 --server.port 8501
```

#### 3. 使用 systemd 管理服务（Linux）

创建服务文件 `/etc/systemd/system/cost-dashboard.service`:

```ini
[Unit]
Description=Cost Dashboard Streamlit App
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/path/to/your/project
ExecStart=/usr/bin/streamlit run main.py --server.address 0.0.0.0 --server.port 8501
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

启动服务：
```bash
sudo systemctl enable cost-dashboard
sudo systemctl start cost-dashboard
sudo systemctl status cost-dashboard
```

## 服务器部署建议

### 1. 使用 Docker 的优势
- ✅ **环境一致性**: 开发、测试、生产环境完全一致
- ✅ **依赖隔离**: 避免与系统环境冲突
- ✅ **版本管理**: 支持回滚和版本控制
- ✅ **资源控制**: 可限制内存和CPU使用
- ✅ **部署简化**: 一条命令完成部署
- ✅ **扩展性**: 支持负载均衡和集群部署

### 2. 生产环境配置

#### 使用 Nginx 反向代理
```nginx
server {
    listen 80;
    server_name your-domain.com;
    
    location / {
        proxy_pass http://localhost:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_cache_bypass $http_upgrade;
    }
}
```

#### 使用 HTTPS（推荐）
```bash
# 使用 Let's Encrypt 免费证书
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d your-domain.com
```

### 3. 监控和日志

#### 日志管理
```bash
# Docker 日志
docker-compose logs -f cost-dashboard

# 系统日志（直接部署）
journalctl -u cost-dashboard -f
```

#### 性能监控
- 使用 `docker stats` 监控容器资源使用
- 配置日志轮转避免磁盘空间不足
- 定期检查应用健康状态

## 安全建议

### 1. 网络安全
- 使用防火墙限制访问端口
- 配置 HTTPS 加密传输
- 使用 VPN 或内网访问

### 2. 应用安全
- 定期更新依赖包
- 限制文件上传大小
- 配置访问控制

### 3. 数据安全
- 定期备份重要数据
- 使用安全的文件传输方式
- 配置数据加密

## 故障排除

### 常见问题

#### 1. 端口被占用
```bash
# 查看端口占用
netstat -tulpn | grep 8501

# 修改端口
# 在 docker-compose.yml 中修改端口映射
ports:
  - "8502:8501"  # 改为其他端口
```

#### 2. 内存不足
```bash
# 增加内存限制
# 在 docker-compose.yml 中修改
deploy:
  resources:
    limits:
      memory: 4G  # 增加内存限制
```

#### 3. 服务无法启动
```bash
# 查看详细日志
docker-compose logs cost-dashboard

# 检查配置文件
docker-compose config
```

## 更新和维护

### 应用更新
```bash
# 拉取最新代码
git pull

# 重新构建并部署
docker-compose up -d --build
```

### 依赖更新
```bash
# 更新 requirements.txt 后重新构建
docker-compose build --no-cache
docker-compose up -d
```

## 联系支持

如遇到部署问题，请提供以下信息：
- 操作系统版本
- Docker 版本
- 错误日志
- 部署方式（Docker/直接部署）