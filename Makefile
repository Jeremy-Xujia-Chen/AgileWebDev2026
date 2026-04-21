# StudySync — common tasks (requires GNU Make; on Windows use Git Bash or WSL).
PYTHON := ./.venv/bin/python
PIP := ./.venv/bin/pip
PYTEST := ./.venv/bin/pytest
FLASK := ./.venv/bin/flask

.PHONY: venv install test smoke run help

help:
	@echo "Targets:"
	@echo "  make venv     - create .venv (if missing) and install requirements"
	@echo "  make install  - same as venv"
	@echo "  make test     - run pytest"
	@echo "  make smoke    - run scripts/smoke_checks.py (no server)"
	@echo "  make run      - flask run on 127.0.0.1:5000"

venv install:
	@bash scripts/setup_venv.sh

test:
	@$(PYTEST) $(ARGS)

smoke:
	@$(PYTHON) scripts/smoke_checks.py

run:
	@FLASK_APP=run.py $(FLASK) run --host 127.0.0.1 --port 5000
