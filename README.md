# StudySync (AgileWebDev2026)

StudySync is a Flask-based web app for **timetable management**, **exam preparation**, **study groups** (merged schedules and shared tasks), and an **AI study planner** (backend stub until the teammate connects the model).

**Stack (unit brief):** Flask, SQLite via **SQLAlchemy**, Bootstrap + **jQuery**, **AJAX** for JSON APIs (no WebSockets required). Forms use **CSRF** (Flask-WTF); passwords stored as **hashes**.

## Purpose and design

See [`docs/PROJECT_IMPLEMENTATION_PLAN.md`](docs/PROJECT_IMPLEMENTATION_PLAN.md) for the full implementation plan, database outline, API list, phased roadmap, and the **AI Planner HTTP contract** (§8).

High-level idea (also in `项目构想.md`): extend a CAS-style weekly timetable with exam prep, collaboration views, and AI-assisted planning.

## Group members

| UWA ID | Name | GitHub |
|--------|------|--------|
| *(add)* | *(add)* | *(add)* |

---

## 启动项目（中文简要）

分三件事理解：**装依赖**、**激活环境**、**用 Python 启动服务**。它们不是同一条命令。

### 1）安装依赖（每个克隆 / 改 `requirements.txt` 后做一次）

`bash scripts/setup_venv.sh` 会做两件：**创建 `.venv`**，并在其中执行 **`pip install -r requirements.txt`**，Flask 等依赖会装进虚拟环境，**不会**污染系统 Python。

等价手动命令：

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows 见下
pip install -r requirements.txt
```

### 2）`source .venv/bin/activate` 是什么意思？

**它不会启动网站，也不会安装包。** 只是让当前终端里打 `python` / `pip` / `flask` 时，优先用 **`.venv` 里的版本**（PATH 前面挂上 `.venv/bin`）。  
**新开一个终端**想继续开发时，需要再执行一次 `source`（或下面用**绝对路径**启动，则可不激活）。

- macOS / Linux: `source .venv/bin/activate`
- Windows CMD: `.venv\Scripts\activate.bat`
- Windows PowerShell: `.venv\Scripts\Activate.ps1`

### 3）真正启动开发服务器（必须再执行一条）

在**已激活** `.venv` 的终端里，任选其一：

```bash
python3 run.py
```

或：

```bash
export FLASK_APP=run.py
flask run --host 127.0.0.1 --port 5000
```

或（脚本内部等价于用 `.venv` 里的 `flask run`）：

```bash
bash scripts/run_dev.sh
```

**不激活**也可以，直接用虚拟环境里的解释器（路径随系统略不同，核心是调用 `.venv/bin/python`）：

```bash
.venv/bin/python3 run.py
```

看到终端里出现 `Running on http://127.0.0.1:5000` 即表示服务已启动。

### 4）浏览器与数据库

打开 **[http://127.0.0.1:5000](http://127.0.0.1:5000)** → 登录/注册 → **[http://127.0.0.1:5000/timetable](http://127.0.0.1:5000/timetable)**。  
数据库默认：**`instance/studysync.db`**（首次请求时创建）。

### 5）环境变量（可选）

`cp .env.example .env`，把 `SECRET_KEY` 改成随机长字符串（本地开发可不改也能跑，**部署前必须改**）。

---

## How to run locally (English)

**Three separate steps:** install deps → (optional) activate venv → start the server.

```bash
bash scripts/setup_venv.sh        # creates .venv AND pip install -r requirements.txt
source .venv/bin/activate         # Windows: .venv\Scripts\activate — only adjusts PATH; does not install or start the app
cp .env.example .env              # optional: set SECRET_KEY
python3 run.py                    # or: export FLASK_APP=run.py && flask run --host 127.0.0.1 --port 5000
# or: bash scripts/run_dev.sh
# without activating: .venv/bin/python3 run.py
```

Useful URLs:

- [http://127.0.0.1:5000/health](http://127.0.0.1:5000/health) — JSON health check  
- [http://127.0.0.1:5000/login](http://127.0.0.1:5000/login) — sign in / register  
- [http://127.0.0.1:5000/timetable](http://127.0.0.1:5000/timetable) — weekly timetable (AJAX) after login  
- `POST /api/planner/chat` with JSON `{"message":"hello"}` — AI stub (no model yet)  
- Timetable API: `GET /api/timetable/events?week_start=YYYY-MM-DD`; mutations under `/api/timetable/events` need header **`X-CSRFToken`** (same as `<meta name="csrf-token">` on the timetable page).

---

## 如何测试（中文简要）

| 方式 | 命令 | 说明 |
|------|------|------|
| **完整自动化测试** | `bash scripts/run_tests.sh` 或 `pytest` | 与作业要求一致，推荐每次提交前跑 |
| **快速冒烟（不启动服务器）** | `python3 scripts/smoke_checks.py` | 用 Flask 内置 test client 走一遍关键接口 |
| **对正在运行的服务做 HTTP 探测** | 终端 A：`bash scripts/run_dev.sh`；终端 B：`bash scripts/curl_smoke.sh` | 只测 `/health` 和 AI stub，不含登录 |

有 **Make** 时：`make venv`、`make test`、`make smoke`、`make run`（见 Makefile）。

Windows 没有 bash 时：用 `python -m venv .venv` + `pip install -r requirements.txt`，测试可用 `scripts\run_tests.bat`。

---

## How to run tests (English)

```bash
source .venv/bin/activate
bash scripts/run_tests.sh           # same as: pytest
bash scripts/run_tests.sh -v      # verbose
bash scripts/run_tests.sh tests/test_timetable.py   # single file
```

The unit rubric expects **5+ unit tests** and **5+ Selenium tests**; Selenium will be added later. Current suite is under [`tests/`](tests/).

---

## Scripts in `scripts/`

| Script | Purpose |
|--------|---------|
| [`scripts/setup_venv.sh`](scripts/setup_venv.sh) | Create `.venv` and `pip install -r requirements.txt` |
| [`scripts/run_dev.sh`](scripts/run_dev.sh) | `flask run` on `127.0.0.1:5000` (set `PORT=8000` to change port) |
| [`scripts/run_tests.sh`](scripts/run_tests.sh) | Run `pytest` (passes extra args through) |
| [`scripts/smoke_checks.py`](scripts/smoke_checks.py) | Fast in-process smoke checks (no browser) |
| [`scripts/curl_smoke.sh`](scripts/curl_smoke.sh) | `curl` `/health` + AI stub against a **running** server |
| [`scripts/run_tests.bat`](scripts/run_tests.bat) | Windows: run pytest |

---

## Prototype HTML (repo root)

The original **static** mockups (`timetable.html`, `ai_planner.html`, `group.html`, `exam_detail.html`, `login.html`) are still in the **repository root** as design references.

- **Timetable:** the Flask page `/timetable` now reuses the same **StudySync** look (sidebar, top bar, filters, dark theme) and wires **real data** via AJAX. You do **not** install those files separately; we **merge** layout/CSS/JS into `templates/` + `static/` over time.
- **Login:** still served from `templates/auth/login_register.html` (simplified styling). Porting the exact `login.html` look is optional polish.
- **AI / Group / Exams:** sidebar links go to `/ai-planner`, `/group`, `/exams` (stub pages) until those screens are migrated from the root HTML files.
