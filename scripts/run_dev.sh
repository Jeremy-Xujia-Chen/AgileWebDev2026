#!/usr/bin/env bash
# Start the Flask development server (http://127.0.0.1:5000 by default).
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ ! -x .venv/bin/flask ]]; then
  echo "Missing .venv or Flask. Run: bash scripts/setup_venv.sh" >&2
  exit 1
fi

export FLASK_APP="${FLASK_APP:-run.py}"
export FLASK_DEBUG="${FLASK_DEBUG:-1}"
echo "Starting Flask (FLASK_APP=$FLASK_APP)…"
exec .venv/bin/flask run --host 127.0.0.1 --port "${PORT:-5000}"
