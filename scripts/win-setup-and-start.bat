@echo off
chcp 65001 >nul
title 招标代理 SOP 工作台 — Windows 一键部署与启动
setlocal enabledelayedexpansion

set "ROOT=%~dp0.."
cd /d "%ROOT%"

echo ==================================================
echo   招标代理 SOP 工作台 — Windows 傻瓜一键启动
echo ==================================================
echo.

:: 1. 检查 Node.js（优先使用内置 node\node.exe，其次系统 PATH）
set "NODE_EXE=%ROOT%\node\node.exe"
set "NPM_CMD=npm"
if exist "%NODE_EXE%" (
  echo [1/4] 检测到内置绿色版 Node.js
) else (
  where node >nul 2>nul
  if errorlevel 1 (
    echo [1/4] 未检测到 Node.js，尝试通过 winget 自动安装 Node.js LTS...
    where winget >nul 2>nul
    if errorlevel 1 (
      echo.
      echo [提示] 当前系统没有 winget，请手动安装 Node.js v22.10+ 或 v24 LTS：
      echo        https://nodejs.org/
      echo        安装完成后重新双击本脚本即可！
      pause
      exit /b 1
    )
    winget install -e --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
    echo.
    echo [提示] Node.js 安装完成后，请关闭本窗口并重新双击本脚本继续。
    pause
    exit /b 0
  )
  set "NODE_EXE=node"
  for /f "delims=" %%v in ('node -v') do set "NODE_VER=%%v"
  echo [1/4] 已检测到系统 Node.js: !NODE_VER!
)

:: 2. 自动安装依赖（仅首次或缺失 node_modules 时执行）
if not exist "%ROOT%\server\node_modules" (
  echo [2/4] 首次运行，正在安装依赖（需要联网，大约 1-2 分钟）...
  call %NPM_CMD% run install:all
  if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接后重试。
    pause
    exit /b 1
  )
) else (
  echo [2/4] 依赖已就绪，跳过安装。
)

:: 3. 自动编译前后端产物（仅缺失 dist 时执行）
if not exist "%ROOT%\server\dist\index.js" (
  echo [3/4] 正在编译前端与后端...
  call %NPM_CMD% run build
  if errorlevel 1 (
    echo [错误] 编译失败，请查看上方报错信息。
    pause
    exit /b 1
  )
) else if not exist "%ROOT%\client\dist\index.html" (
  echo [3/4] 正在编译前端...
  call %NPM_CMD% --prefix client run build
  if errorlevel 1 (
    echo [错误] 前端编译失败，请查看上方报错信息。
    pause
    exit /b 1
  )
) else (
  echo [3/4] 编译产物已就绪，跳过编译。
)

:: 4. 启动服务并自动打开浏览器
set "PORT=3001"
set "AUTH_ENABLED=false"
if "%DATA_DIR%"=="" set "DATA_DIR=%ROOT%\server\data"

echo [4/4] 正在启动工作台服务（端口 %PORT%）...
echo       数据目录：%DATA_DIR%
echo       访问地址：http://localhost:%PORT%
echo.
echo [提示] 浏览器将在 2 秒后自动打开。关闭本窗口即可停止服务。
echo.

start "" /b cmd /c "timeout /t 2 /nobreak >nul && start "" http://localhost:%PORT%"
cd /d "%ROOT%\server"
"%NODE_EXE%" "%ROOT%\server\dist\index.js"
pause
