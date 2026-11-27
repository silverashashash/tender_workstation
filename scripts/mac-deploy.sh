#!/usr/bin/env bash
set -euo pipefail

# Mac mini 方案一：一键编译并注册/重启 macOS 后台常驻服务 (launchd)
# 用法：
#   npm run mac:deploy          # 本地编译并重启后台服务
#   npm run mac:deploy -- --pull # 先 git pull 最新代码，再编译并重启

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PLIST_LABEL="com.tender.workstation"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"
LOG_DIR="$ROOT_DIR/server/data"
mkdir -p "$HOME/Library/LaunchAgents" "$LOG_DIR"

# 自动定位 node 路径（兼容 Homebrew / nvm / 系统默认）
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
if [ -d "$HOME/.nvm/versions/node" ]; then
  LATEST_NVM="$(ls "$HOME/.nvm/versions/node" | sort -V | tail -n 1)"
  if [ -n "$LATEST_NVM" ]; then
    export PATH="$HOME/.nvm/versions/node/$LATEST_NVM/bin:$PATH"
  fi
fi

NODE_BIN="$(command -v node || true)"
NPM_BIN="$(command -v npm || true)"
if [ -z "$NODE_BIN" ] || [ -z "$NPM_BIN" ]; then
  echo "[错误] 未找到 node 或 npm，请先安装 Node.js (v22.10+)。"
  exit 1
fi

echo "=== 招标代理 SOP 工作台 · Mac mini 后台部署 ==="
echo "项目目录: $ROOT_DIR"
echo "Node 路径: $NODE_BIN ($("$NODE_BIN" -v))"

cd "$ROOT_DIR"

if [[ "${1:-}" == "--pull" ]]; then
  echo "[1/3] 拉取 GitHub 最新代码..."
  git pull --rebase origin main
else
  echo "[1/3] 使用当前本地代码（如需先拉取远程，可运行 npm run mac:deploy -- --pull）"
fi

echo "[2/3] 编译前端与后端产物..."
"$NPM_BIN" run build

echo "[3/3] 写入并重启 launchd 后台常驻服务 ($PLIST_LABEL)..."
cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${PLIST_LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${NODE_BIN}</string>
    <string>${ROOT_DIR}/server/dist/index.js</string>
  </array>
  <key>WorkingDirectory</key>
  <string>${ROOT_DIR}/server</string>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${LOG_DIR}/server.log</string>
  <key>StandardErrorPath</key>
  <string>${LOG_DIR}/server.err.log</string>
</dict>
</plist>
EOF

UID_NUM="$(id -u)"
launchctl bootout "gui/${UID_NUM}" "$PLIST_PATH" 2>/dev/null || true
launchctl bootstrap "gui/${UID_NUM}" "$PLIST_PATH"
launchctl kickstart -k "gui/${UID_NUM}/${PLIST_LABEL}"

sleep 1
PORT_VAL="3001"
if [ -f "$ROOT_DIR/.env" ]; then
  ENV_PORT="$(grep -E '^PORT=' "$ROOT_DIR/.env" | tail -n 1 | cut -d'=' -f2 | tr -d '[:space:]' || true)"
  if [ -n "$ENV_PORT" ]; then PORT_VAL="$ENV_PORT"; fi
fi

echo ""
echo "[4/4] 同步刷新 Windows 绿色便携版产物 (release/tender-workstation-win-portable)..."
"$NODE_BIN" "$ROOT_DIR/scripts/build-win-portable.mjs" --skip-build

echo ""
echo "✅ 后台服务已启动并设置为开机自启，Windows 便携版也已同步更新！"
echo "   本机访问：http://localhost:${PORT_VAL}"
TS_IP="$(tailscale ip -4 2>/dev/null | head -n 1 || true)"
if [ -n "$TS_IP" ]; then
  echo "   MacBook 通过 Tailscale 访问：http://${TS_IP}:${PORT_VAL}"
else
  echo "   MacBook 通过 Tailscale 访问：http://<你的Mac-mini-Tailscale-IP>:${PORT_VAL}"
fi
echo "   Windows 便携版目录：${ROOT_DIR}/release/tender-workstation-win-portable"
echo "   运行日志：${LOG_DIR}/server.log"
echo "   停止服务命令：launchctl bootout gui/${UID_NUM} ${PLIST_PATH}"
