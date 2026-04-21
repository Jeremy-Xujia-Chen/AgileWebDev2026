#!/usr/bin/env bash
# Quick HTTP smoke checks against a *running* server (default http://127.0.0.1:5000).
# Usage: in one terminal: bash scripts/run_dev.sh
#         in another:     bash scripts/curl_smoke.sh
set -euo pipefail
BASE="${1:-http://127.0.0.1:5000}"

echo "GET $BASE/health"
curl -sfS "$BASE/health" | head -c 200
echo
echo "OK: /health"

echo "POST $BASE/api/planner/chat (stub)"
curl -sfS -X POST "$BASE/api/planner/chat" \
  -H "Content-Type: application/json" \
  -d '{"message":"ping"}' | head -c 300
echo
echo "OK: AI stub"

echo "Done. (Login + timetable require browser or pytest.)"
