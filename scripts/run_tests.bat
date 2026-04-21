@echo off
REM Run pytest on Windows (from repo root, after setup_venv equivalent).
setlocal
cd /d "%~dp0.."
if not exist ".venv\Scripts\pytest.exe" (
  echo Missing .venv. Create it: python -m venv .venv ^& .venv\Scripts\pip install -r requirements.txt
  exit /b 1
)
.venv\Scripts\pytest.exe %*
