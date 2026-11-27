import { execFileSync, execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, '..');
const RELEASE_ROOT = path.join(ROOT_DIR, 'release');
const OUT_DIR = path.join(RELEASE_ROOT, 'tender-workstation-win-portable');
const WIN_NODE_CACHE = path.join(RELEASE_ROOT, 'win-node-cache');

const includeData = process.argv.includes('--include-data');
const skipBuild = process.argv.includes('--skip-build');

function copyDir(src, dest) {
  if (!fs.existsSync(src)) return;
  fs.mkdirSync(dest, { recursive: true });
  fs.cpSync(src, dest, { recursive: true });
}

/**
 * Extract a Node.js Windows zip (e.g. node-v26.13.1-win-x64.zip) into destDir,
 * flattening the inner single root folder so destDir/node.exe exists directly.
 */
function extractNodeZip(zipPath, destDir) {
  const tmpDir = `${destDir}.__tmp__`;
  fs.rmSync(tmpDir, { recursive: true, force: true });
  fs.rmSync(destDir, { recursive: true, force: true });
  fs.mkdirSync(tmpDir, { recursive: true });
  execFileSync('unzip', ['-q', zipPath, '-d', tmpDir], { stdio: 'inherit' });

  const entries = fs.readdirSync(tmpDir).filter((n) => !n.startsWith('__MACOSX') && !n.startsWith('.'));
  if (entries.length === 1 && fs.statSync(path.join(tmpDir, entries[0])).isDirectory()) {
    fs.renameSync(path.join(tmpDir, entries[0]), destDir);
    fs.rmSync(tmpDir, { recursive: true, force: true });
  } else {
    fs.renameSync(tmpDir, destDir);
  }
}

console.log('=== 开始构建 Windows 绿色便携版（双架构自动适配）===');
if (!skipBuild) {
  console.log('[1/4] 编译前端与后端产物...');
  execSync('npm run build', { cwd: ROOT_DIR, stdio: 'inherit' });
} else {
  console.log('[1/4] 跳过编译（--skip-build）...');
}

console.log('[2/4] 清理并组装便携目录:', OUT_DIR);
fs.rmSync(OUT_DIR, { recursive: true, force: true });
fs.mkdirSync(OUT_DIR, { recursive: true });

copyDir(path.join(ROOT_DIR, 'client', 'dist'), path.join(OUT_DIR, 'client', 'dist'));
copyDir(path.join(ROOT_DIR, 'server', 'dist'), path.join(OUT_DIR, 'server', 'dist'));
copyDir(path.join(ROOT_DIR, 'server', 'templates'), path.join(OUT_DIR, 'server', 'templates'));
fs.copyFileSync(
  path.join(ROOT_DIR, 'server', 'package.json'),
  path.join(OUT_DIR, 'server', 'package.json')
);
copyDir(path.join(ROOT_DIR, 'server', 'node_modules'), path.join(OUT_DIR, 'server', 'node_modules'));

const portableDataDir = path.join(OUT_DIR, 'portable_data');
fs.mkdirSync(path.join(portableDataDir, 'uploads'), { recursive: true });

if (includeData) {
  console.log('[3/4] 打包当前本地数据库与上传附件（--include-data）...');
  const srcDb = path.join(ROOT_DIR, 'server', 'data', 'workstation.db');
  const srcUploads = path.join(ROOT_DIR, 'server', 'uploads');
  if (fs.existsSync(srcDb)) {
    fs.copyFileSync(srcDb, path.join(portableDataDir, 'workstation.db'));
  }
  if (fs.existsSync(srcUploads)) {
    copyDir(srcUploads, path.join(portableDataDir, 'uploads'));
  }
} else {
  console.log('[3/4] 创建空白 portable_data 目录（如需附带当前本机数据，可加参数 --include-data）');
}

// Auto-embed Windows portable Node.js from release/win-node-cache/
// Supports BOTH x64 and arm64 zips simultaneously!
const bundledArchs = [];
if (fs.existsSync(WIN_NODE_CACHE)) {
  const cacheItems = fs.readdirSync(WIN_NODE_CACHE);
  const x64Zip = cacheItems.find((n) => /win-x64.*\.zip$/i.test(n));
  const arm64Zip = cacheItems.find((n) => /win-arm64.*\.zip$/i.test(n));

  if (x64Zip) {
    console.log('      正在自动内置 Intel/AMD (x64) Node.js:', x64Zip);
    extractNodeZip(path.join(WIN_NODE_CACHE, x64Zip), path.join(OUT_DIR, 'node-x64'));
    bundledArchs.push('x64');
  }
  if (arm64Zip) {
    console.log('      正在自动内置 ARM64 Node.js:', arm64Zip);
    extractNodeZip(path.join(WIN_NODE_CACHE, arm64Zip), path.join(OUT_DIR, 'node-arm64'));
    bundledArchs.push('arm64');
  }
  if (bundledArchs.length === 0 && fs.existsSync(path.join(WIN_NODE_CACHE, 'node.exe'))) {
    copyDir(WIN_NODE_CACHE, path.join(OUT_DIR, 'node'));
    bundledArchs.push('generic');
  }
}

if (bundledArchs.length === 0) {
  const nodeDir = path.join(OUT_DIR, 'node');
  fs.mkdirSync(nodeDir, { recursive: true });
  fs.writeFileSync(
    path.join(nodeDir, '请将Windows绿色版Node解压至此目录.txt'),
    [
      '如果目标 Windows 电脑未安装 Node.js (v22.10+)，请执行以下一次性准备：',
      '1. 前往 https://nodejs.org/dist/ 下载 Windows 绿色压缩包（Intel/AMD 用 win-x64.zip，ARM 用 win-arm64.zip）。',
      '2. 直接把 zip 文件放进项目根目录的 release/win-node-cache/，重新运行 npm run build:win-portable 即可自动内置双架构！',
    ].join('\r\n'),
    'utf8'
  );
} else {
  console.log('      已内置 Windows Node.js 架构:', bundledArchs.join(' + '));
}

console.log('[4/4] 写入双架构自动识别启动脚本与使用说明...');
const startBat = [
  '@echo off',
  'chcp 65001 >nul',
  'title 招标代理 SOP 工作台 (内网便携版)',
  'setlocal',
  'set "ROOT=%~dp0"',
  'set "PORT=3001"',
  '',
  ':: 检查 3001 端口是否已在运行，若已运行则直接打开浏览器',
  'netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul 2>nul',
  'if not errorlevel 1 (',
  '  echo 检测到工作台已在运行，正在直接打开浏览器...',
  '  start "" http://localhost:%PORT%',
  '  timeout /t 2 /nobreak >nul',
  '  exit /b 0',
  ')',
  '',
  ':: 自动识别 Windows CPU 架构 (ARM64 vs x64/Intel/AMD) 并选用对应内置 Node.js',
  'set "ARCH=%PROCESSOR_ARCHITECTURE%"',
  'if defined PROCESSOR_ARCHITEW6432 set "ARCH=%PROCESSOR_ARCHITEW6432%"',
  'set "NODE_EXE="',
  'if /I "%ARCH%"=="ARM64" (',
  '  if exist "%ROOT%node-arm64\\node.exe" set "NODE_EXE=%ROOT%node-arm64\\node.exe"',
  ') else (',
  '  if exist "%ROOT%node-x64\\node.exe" set "NODE_EXE=%ROOT%node-x64\\node.exe"',
  ')',
  'if not defined NODE_EXE if exist "%ROOT%node\\node.exe" set "NODE_EXE=%ROOT%node\\node.exe"',
  'if not defined NODE_EXE (',
  '  where node >nul 2>nul',
  '  if errorlevel 1 (',
  '    echo [错误] 未检测到适配当前架构 ^(%ARCH%^) 的 Node.js！',
  '    echo 请将对应架构的 Windows 绿色版 Node.js ^(v22.10+^) 放入 node-x64 或 node-arm64 文件夹。',
  '    pause',
  '    exit /b 1',
  '  )',
  '  set "NODE_EXE=node"',
  ')',
  '',
  'set "AUTH_ENABLED=false"',
  'set "DATA_DIR=%ROOT%portable_data"',
  'echo ==================================================',
  'echo   招标代理 SOP 工作台 (Windows 双架构内网便携版)',
  'echo ==================================================',
  'echo 当前系统架构: %ARCH%',
  'echo 使用 Node:    %NODE_EXE%',
  'echo 数据目录:     %DATA_DIR%',
  'echo 访问地址:     http://localhost:%PORT%',
  'echo.',
  'echo [提示] 浏览器将在 2 秒后自动打开。',
  'echo [离场] 请先在页面左下角点「导出当前批次 (.zip)」，再关闭本窗口或双击「停止工作台.bat」后拔 U 盘！',
  'echo.',
  'start "" /b cmd /c "timeout /t 2 /nobreak >nul && start "" http://localhost:%PORT%"',
  'cd /d "%ROOT%server"',
  '"%NODE_EXE%" "%ROOT%server\\dist\\index.js"',
  'pause',
].join('\r\n');

fs.writeFileSync(path.join(OUT_DIR, '启动工作台.bat'), startBat, 'utf8');

const stopBat = [
  '@echo off',
  'chcp 65001 >nul',
  'title 停止招标代理 SOP 工作台',
  'set "PORT=3001"',
  'echo 正在安全停止端口 %PORT% 上的工作台服务...',
  'for /f "tokens=5" %%a in (\'netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"\') do (',
  '  taskkill /PID %%a /F >nul 2>nul',
  ')',
  'echo 已停止工作台服务，等待 2 秒后即可安全弹出 U 盘。',
  'timeout /t 2 /nobreak >nul',
].join('\r\n');

fs.writeFileSync(path.join(OUT_DIR, '停止工作台.bat'), stopBat, 'utf8');

const readmeTxt = [
  '==================================================',
  '  招标代理 SOP 工作台 — Windows 双架构内网绿色便携版',
  '==================================================',
  '',
  '【一、免安装自动识别架构】',
  '本便携包已支持同时内置 Intel/AMD (node-x64) 与 ARM64 (node-arm64) 双版本 Node.js，',
  '双击「启动工作台.bat」时会自动检测当前 Windows 电脑架构并选用对应内核，无需手动切换。',
  '',
  '【二、在 Windows 电脑 / Parallels 虚拟机中运行】',
  '1. 把整个文件夹拷到 Windows 本地盘（如桌面；Parallels 中请勿直接在共享盘运行），双击「启动工作台.bat」。',
  '2. 2 秒后浏览器会自动打开 http://localhost:3001。',
  '3. 全新机器初始化：点左下角「恢复整库备份 (.zip)」一键同步 Mac 全部历史数据；',
  '   或只带单个项目：点左下角「导入批次包 (.zip)」载入单批次数据。',
  '',
  '【三、开标结束后带回 Mac mini】',
  '1. 在页面左下角点击「导出当前批次 (.zip)」，保存单个 ZIP 文件带出。',
  '2. 关闭黑色命令行窗口（或双击「停止工作台.bat」），等待 2 秒再拔 U 盘。',
  '3. 回到 Mac 浏览器，点左下角「导入批次包 (.zip)」即可无损合并更新，不影响其他批次！',
].join('\r\n');

fs.writeFileSync(path.join(OUT_DIR, '使用说明.txt'), readmeTxt, 'utf8');

console.log('=== 打包完成 ===');
console.log('输出目录:', OUT_DIR);
