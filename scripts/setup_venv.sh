#!/usr/bin/env bash
# Create .venv AND install dependencies: pip install -r requirements.txt
# (Does NOT start the Flask server — run python3 run.py or scripts/run_dev.sh after: source .venv/bin/activate)
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
  echo "Created .venv"
fi

./.venv/bin/pip install -U pip
./.venv/bin/pip install -r requirements.txt
echo "Dependencies installed. Activate with: source .venv/bin/activate"
