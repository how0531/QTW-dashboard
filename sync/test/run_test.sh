#!/usr/bin/env bash
# 整條流程測試：假的 C:\XS\data → sheet_sync.py → 假的 Apps Script（跑真的 Code.gs）→ 檢查試算表內容
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
WORK="$(mktemp -d)"
PORT=18765
TOKEN=test-token-123
# 假試算表每頁只給 30 列，用少量資料就能測到「列數不夠要自己加列」
FAKE_MAX_ROWS=30 node "$HERE/fake_gas_server.js" $PORT $TOKEN > "$WORK/server.log" 2>&1 &
SRV=$!
trap 'kill $SRV 2>/dev/null; rm -rf "$WORK"' EXIT
for i in $(seq 1 50); do grep -q ready "$WORK/server.log" 2>/dev/null && break; sleep 0.1; done
python3 -I "$HERE/check.py" "$HERE/.." "$WORK" "http://127.0.0.1:$PORT" "$TOKEN"
