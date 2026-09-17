#!/usr/bin/env bash
# =============================================================================
# PO Cutting — VPS Deploy Script
# Run this ON the VPS (e.g. `bash deploy-vps.sh` inside the app directory).
#
# What it does:
#   1. git pull (fast-forward only, from madison88admin/PO_Cutting_Automation)
#   2. npm ci (clean install — package.json changed: jsonrepair, ssh2, fuse.js)
#   3. npm run build
#   4. restart the app (pm2 if present, otherwise systemd, otherwise npm start)
#   5. verify /api/health and the new /api/nextgen-po-lines route
#
# Environment file: make sure .env / .env.local on the VPS has at least:
#   SUPABASE_* credentials (po_cutting schema)
#   NEXTGEN_* read credentials
#   NEXTGEN_WRITE_* (only if PO insert is used)
#   NEXTGEN_REQUEST_TIMEOUT_MS=45000   <- recommended for big PO pulls (Power BI)
#   PO_LINE_DUMP=/path/to/PO Line Data Dump.xlsx   <- for /api/diff-report
# =============================================================================
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$APP_DIR"
echo "==> App dir: $APP_DIR"

echo "==> Current commit before pull:"
git log --oneline -1

# 1. Pull (never rebase/force — fast-forward only for safety)
echo "==> Pulling latest from origin..."
git pull --ff-only origin main

echo "==> New commit:"
git log --oneline -1

# 2. Clean install (package.json changed in this batch)
echo "==> Installing dependencies (npm ci)..."
npm ci

# 3. Build
echo "==> Building..."
npm run build

# 4. Restart — detect process manager
if command -v pm2 >/dev/null 2>&1 && pm2 list 2>/dev/null | grep -qi "po"; then
    echo "==> Restarting via pm2..."
    pm2 restart po-cutting || pm2 restart all
elif systemctl list-units --type=service 2>/dev/null | grep -qi "po-cutting\|pocutting\|nextjs"; then
    SVC="$(systemctl list-units --type=service --no-legend | grep -i 'po-cutting\|pocutting\|nextjs' | awk '{print $1}' | head -1)"
    echo "==> Restarting systemd service: $SVC (needs sudo)"
    sudo systemctl restart "$SVC"
else
    echo "!! No pm2/systemd unit detected."
    echo "!! Start manually with:  nohup npm start > /var/log/po-cutting.log 2>&1 &"
    echo "!! (or tell your agent how the app is supervised on this box)"
fi

# 5. Verify
echo "==> Waiting for app to come up..."
sleep 6
PORT="${PORT:-3000}"
echo -n "health:  "; curl -s -o /dev/null -w "%{http_code}\n" "http://localhost:${PORT}/api/health" || true
echo -n "po-lines: "; curl -s -o /dev/null -w "%{http_code}\n" "http://localhost:${PORT}/api/nextgen-po-lines?pageSize=1" || true

echo ""
echo "==> If both are 200, verify publicly:"
echo "    curl 'https://po-cutting.5-223-78-194.sslip.io/api/nextgen-po-lines?page=1&pageSize=1'"
