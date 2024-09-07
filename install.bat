@echo off
setlocal enabledelayedexpansion

:: 检查Docker是否已安装
docker --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Docker Desktop is not installed. Installing...
    :: 下载 Docker Desktop
    start /wait DockerInstaller.exe install
    echo Docker has been successfully installed.
    pause
    exit /b
)

:: 启动Docker Desktop
echo Starting Docker Desktop...
start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"
timeout /t 10

:: 等待 Docker Desktop 启动
:waitfordocker
docker info >nul 2>&1
if %errorlevel% neq 0 (
    echo Waiting for Docker Desktop to start...
    timeout /t 5
    goto waitfordocker
)

:: 检查并导入后端镜像
docker images | findstr " excel-app " >nul 2>&1
if %errorlevel% neq 0 (
    echo "excel-app" image not found, loading the image...
    docker load -i .\excel-app-image.tar
) else (
    echo "excel-app" image already exists, skipping load.
)

:: 检查并导入前端镜像
docker images | findstr " excel-app-portal " >nul 2>&1
if %errorlevel% neq 0 (
    echo "excel-app-portal" image not found, loading the image...
    docker load -i .\excel-app-portal-image.tar
) else (
    echo "excel-app-portal" image already exists, skipping load.
)

:: 运行 docker-compose 启动容器
echo Starting container...
docker-compose up -d

:: 完成
echo Service started, access frontend address: http://localhost:8081
pause
