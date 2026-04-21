#!/usr/bin/env bash
# Run the automated test suite (pytest). Pass extra args through, e.g. -v or a file path.
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if [[ ! -x .venv/bin/pytest ]]; then
  echo "Missing .venv or pytest. Run: bash scripts/setup_venv.sh" >&2
  exit 1
fi

exec .venv/bin/pytest "$@"
