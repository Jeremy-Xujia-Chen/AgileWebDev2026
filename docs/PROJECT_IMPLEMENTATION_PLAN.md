# StudySync 详细项目实现方案

本文档与课程作业（CITS3403 / CITS5505）要求对齐，约定技术栈、数据模型、接口与分阶段交付。**AI Planner 的后端调用与外部模型集成由组员后续实现**；当前仅保留**明确的 HTTP 契约与占位实现**，其余功能按阶段逐步实现。

---

## 1. 目标与范围

### 1.1 产品目标（与 `项目构想.md` 一致）

- 在 **周课表** 思路上管理讲授课、实验、辅导、考试、作业截止等活动。
- **考试 / 任务备考**：进度、清单、笔记、资源链接；可选展示组员备考概况。
- **小组协作**：合并课表、公共空闲时间、小组任务；满足作业「用户可查看其他用户数据」。
- **AI 学习规划**：对话式规划、结构化计划卡片；**接口先占位**，见第 8 节。

### 1.2 作业必须满足的技术与质量项

| 类别 | 要求 |
|------|------|
| 架构 | 客户端 + Flask 服务端；持久化 SQLite + SQLAlchemy |
| 账户 | 登录、登出；密码加盐哈希；CSRF 防护 |
| 数据 | 用户数据跨会话持久化；存在「查看其他用户数据」能力（小组视图等） |
| 前端 | HTML / CSS / JS；Bootstrap；**jQuery**；**AJAX**（或 WebSockets） |
| 测试 | ≥5 单元测试 + ≥5 Selenium（可针对已部署的本地服务） |
| 仓库 | README：目的与设计、组员表、启动方式、测试运行方式 |

---

## 2. 推荐仓库目录结构（目标态）

```
AgileWebDev2026/
├── app/
│   ├── __init__.py              # create_app，注册扩展与蓝图
│   ├── config.py                # 配置类，SECRET_KEY 等从环境变量读取
│   ├── extensions.py            # db, login_manager 等
│   ├── models.py                # SQLAlchemy 模型（或拆分为 models/）
│   ├── forms.py                 # WTForms（若课程允许；否则手写校验 + CSRF）
│   ├── auth/                    # 注册、登录、登出
│   ├── timetable/               # 课表页面与 JSON API
│   ├── exams/                   # 考试/任务、复习项、笔记、资源
│   ├── groups/                  # 小组、成员、合并视图、空闲时间、小组任务
│   ├── reminders/               #（可选）提醒列表与 CRUD
│   ├── preferences/             #（可选）用户偏好
│   └── ai_planner/              # 仅占位路由 + 契约文档；逻辑由组员填充
├── templates/                   # Jinja2：从现有 *.html 逐步迁入
├── static/                      # CSS/JS/图片；现有页面样式可拆分至此
├── tests/
│   ├── conftest.py              # Flask test client、临时 DB
│   ├── unit/                    # ≥5
│   └── selenium/                # ≥5
├── instance/                    # 本地 db.sqlite（gitignore）
├── docs/
│   └── PROJECT_IMPLEMENTATION_PLAN.md
├── requirements.txt
├── .env.example
├── run.py                       # 开发入口：create_app()
└── README.md
```

**迁移策略**：短期可将现有根目录 `*.html` 复制到 `templates/` 并改为 `{% extends "base.html" %}`，链接改为 `url_for`；静态资源引用 `{{ url_for('static', filename='...') }}`。不必一次改完，按阶段替换。

---

## 3. 数据库设计（第一版可实现的表）

以下字段名为建议，实现时可微调；核心是 **关系清晰、便于查询周课表与小组**。

### 3.1 `users`

| 字段 | 说明 |
|------|------|
| id | PK |
| email | 唯一，登录用 |
| password_hash | 加盐哈希（Werkzeug `generate_password_hash`） |
| full_name | 显示名 |
| student_id | 可选，学号展示 |
| created_at | 时间戳 |

### 3.2 `groups`（学习小组，非课程分配小组）

| 字段 | 说明 |
|------|------|
| id | PK |
| name | 如 "CITS3403 Study Group" |
| join_code | 唯一短码，用于邀请加入 |
| created_by_user_id | FK → users |

### 3.3 `group_members`

| 字段 | 说明 |
|------|------|
| id | PK |
| group_id | FK |
| user_id | FK |
| role | `owner` / `member`（可选） |
| joined_at | 时间戳 |

唯一约束：`(group_id, user_id)`。

### 3.4 `courses`（用户本学期课程，可选但利于导航）

| 字段 | 说明 |
|------|------|
| id | PK |
| user_id | FK |
| code | 如 CITS3403 |
| title | 课程名 |

### 3.5 `calendar_events`（课表事件）

| 字段 | 说明 |
|------|------|
| id | PK |
| user_id | FK |
| course_id | 可空 FK |
| title | 显示标题 |
| event_type | lecture / lab / tutorial / exam / assignment / workshop / other |
| start_at, end_at | UTC 或本地时区一致即可（全项目统一） |
| location | 可空 |
| notes | 可空 |
| created_at, updated_at | |

列表视图、筛选、周导航均对该表查询。

### 3.6 `exam_sessions`（备考「会话」/重大评估）

可与 `calendar_events` 中 `event_type=exam` 关联，或独立表再 **可选** `calendar_event_id` 外键。

| 字段 | 说明 |
|------|------|
| id | PK |
| user_id | FK |
| course_id | 可空 |
| title | 如 Mid-semester Test |
| starts_at, ends_at | |
| location | 可空 |
| weight_percent | 可空 |
| share_token | 可空，只读分享给他人（可选） |
| created_at | |

### 3.7 `revision_topics`

| 字段 | 说明 |
|------|------|
| id | PK |
| exam_session_id | FK |
| label | 主题文案 |
| sort_order | 整数 |
| status | not_started / in_progress / done |
| progress_percent | 0–100，可与 status 同步 |

### 3.8 `exam_notes`

| 字段 | 说明 |
|------|------|
| id | PK |
| exam_session_id | FK |
| user_id | FK |
| body | 文本 |

一人一笔记即可（与 exam 一对一或按 user+exam 唯一）。

### 3.9 `exam_resources`

| 字段 | 说明 |
|------|------|
| id | PK |
| exam_session_id | FK |
| title, url, resource_type | pdf / video / link 等 |

### 3.10 `group_tasks`

| 字段 | 说明 |
|------|------|
| id | PK |
| group_id | FK |
| title | |
| assignee_user_id | 可空 |
| due_date | 可空 |
| status | not_started / in_progress / done |
| created_by_user_id | FK |

### 3.11 `ai_conversations` / `ai_messages`（可选，组员实现 AI 时再接）

第一版可不做持久化，仅内存会话；若要做「History」，再建表。占位接口应允许后续无缝接表。

---

## 4. 路由与 API 规划

### 4.1 页面路由（HTML，登录保护）

| 路径 | 说明 |
|------|------|
| `/` | 已登录 → 重定向课表；未登录 → 登录 |
| `/login`, `/register` | 登录注册（可同页 Tab） |
| `/logout` | POST 登出 |
| `/timetable` | 周课表 |
| `/exams` | 考试/任务列表 |
| `/exams/<id>` | 备考详情（对应 exam_detail） |
| `/group` 或 `/groups/<id>` | 小组页（对应 group） |
| `/ai-planner` | AI 对话页；前端 AJAX 调占位 API |
| `/courses`, `/reminders`, `/preferences` | 可按里程碑逐步上线 |

### 4.2 JSON API（AJAX，需 CSRF 或 token）

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/timetable/events` | query: `week_start` ISO 日期；返回该周事件列表 |
| POST | `/api/timetable/events` | 创建 |
| PATCH | `/api/timetable/events/<id>` | 更新 |
| DELETE | `/api/timetable/events/<id>` | 删除 |
| GET/PATCH | `/api/exams/<id>` | 详情与元数据更新 |
| GET/POST/PATCH/DELETE | `/api/exams/<id>/topics` | 复习主题 |
| GET/PUT | `/api/exams/<id>/notes` | 笔记保存（可用 debounce 前端） |
| CRUD | `/api/exams/<id>/resources` | 资源 |
| GET | `/api/groups/current` 或 `/api/groups/<id>` | 小组信息 |
| POST | `/api/groups` | 创建小组 |
| POST | `/api/groups/join` | body: `join_code` |
| GET | `/api/groups/<id>/merged-timetable` | 合并课表数据 |
| GET | `/api/groups/<id>/free-slots` | 公共空闲（服务端计算） |
| POST | `/api/groups/<id>/tasks` | 小组任务 |
| PATCH | `/api/groups/<id>/tasks/<task_id>` | 更新状态/指派 |

**AI 占位**见第 8 节，不占用本节实现。

---

## 5. 公共空闲时间算法（建议）

输入：某 `group_id`、一周 `week_start`、步长（如 30 分钟）、参与成员集合（默认全组成员）。

1. 查询该周每个成员的 `calendar_events` 区间。
2. 在周一至周五（或含周末，可配置）的 8:00–20:00 内生成时间槽。
3. 对每个槽判断是否 **所有成员** 在该段均无事件重叠。
4. 合并连续空闲槽为区间，返回 JSON：`[{ "start": "...", "end": "...", "member_count": N }]`

「Book」按钮：创建一个 `group_meeting` 类型事件或写入 `calendar_events` 并标记 `group_id`（若扩展表），便于所有人看到——实现细节可在 Phase 4 定稿。

---

## 6. 分阶段路线图（按顺序做）

### Phase 0 — 工程骨架（第 1 步）

- [x] `requirements.txt`、`run.py`、`create_app`、配置从环境变量读取（含 `python-dotenv`）。
- [x] 注册 **AI Planner 占位蓝图**（`POST /api/planner/chat`，见第 8 节；实现文件 `app/blueprints/ai_planner_stub.py`）。
- [x] `pytest.ini`（`pythonpath = .`）与最小冒烟测试 `tests/test_smoke.py`。
- [x] `templates/` + `static/` 目录与 StudySync 主题；README 随仓库更新（组员表仍待填）。
- **验收**：`flask run` 可访问 `/health` 与占位 AI；`pytest` 通过。

### Phase 1 — 用户认证与 CSRF

- [x] `User` 模型；SQLite + `db.create_all()`；注册/登录/登出；会话 cookie（Flask-Login）。
- [x] **CSRF**：`CSRFProtect` 默认开启；登录/注册表单 `hidden_tag()`；登出 POST 带 `csrf_token()`。
- [x] 密码 **Werkzeug 哈希**；`TestConfig` 中 `WTF_CSRF_ENABLED=False` 便于自动化测试。
- [x] 占位课表页 `GET /timetable` + **AJAX 示例** `GET /api/auth/me`（jQuery）。
- **验收**：注册后刷新仍存在；登出后无法访问受保护页；`pytest` 含认证流测试。

**通信方式**：作业允许「AJAX / WebSockets 二选一或组合」；本项目以 **AJAX（`$.ajax` / `$.getJSON`）** 为主即可，无需 WebSockets。后续 `POST /api/timetable/...` 等需在请求头携带 CSRF（见 Flask-WTF 文档 `X-CSRFToken`）。

### Phase 2 — 个人课表 CRUD

- [x] `calendar_events` 模型（`CalendarEvent`）+ `db.create_all()`。
- [x] 周范围 API：`GET /api/timetable/events?week_start=YYYY-MM-DD`（自动归一化到当周周一；周一至周五重叠查询）。
- [x] `POST /api/timetable/events`、`PATCH/DELETE /api/timetable/events/<id>`；未登录访问 `/api/*` 返回 **401 JSON**（`unauthorized_handler`）。
- [x] `/timetable` 页面：jQuery `$.ajaxSetup` 发送 **`X-CSRFToken`**；周切换、五列日视图、Bootstrap Modal 增删改。
- [x] 类型筛选条；[ ] 与 `source_pages/timetable.html` 或列表视图完全 1:1 对齐（可选）。
- **验收**：多用户数据隔离；他人无法改删你的事件；`pytest` 含课表 API 测试。

### Phase 3 — 考试与备考

- [x] `exam_sessions`、`revision_topics`；考试说明字段在 `ExamSession.notes`；**未**单独建 `exam_notes` / `exam_resources` 表（见 `docs/SOURCE_PAGES_GAP_ANALYSIS.md`）。
- [x] 列表页 + 详情页 + JSON API，与 `source_pages/exam_detail.html` 核心流程对齐；可选增强见 gap 文档。
- [x] 清单（主题）与元数据、笔记类字段的保存（AJAX）。
- [ ]（可选）`share_token` 只读页、独立资源表、按主题细颗粒笔记。
- **验收**：刷新页面数据仍在；权限正确。

### Phase 4 — 小组与「查看他人数据」

- [x] 创建小组、加入码、加入/离开、成员列表。
- [x] 合并课表 API + 页面；公共空闲 `GET` + 展示。
- [ ] UI「**Book**」在空闲段创建共享日历块（见实现方案 §5 与 `SOURCE_PAGES_GAP_ANALYSIS.md`）。
- [x] `group_tasks` API + 页面列表；[ ] 更细权限/指派展示按需。
- **验收**：同组成员可看见合并与任务；他人不可改你私有课表项。

### Phase 5 — 课程、提醒、偏好（按需裁剪）

- [ ] `courses` 与课表/考试关联。
- [ ] `reminders` 与应用内通知列表。
- [ ] `preferences`：时区、默认周起始日等。

### Phase 6 — 测试、安全自查、README

- [x] 单元 / API 测试集（`tests/test_*.py`），含 auth、课表、考试、小组等（≥5）。
- [x] Selenium E2E（`tests/selenium/`，7 条）：健康检查、注册、登录、课表创建、考试导航、**创建小组**、登出。需本机 **Chrome**；`TestConfig` 中 SQLite 使用 `check_same_thread=False` 以配合线程内 E2E 服务。
- [ ] 进一步自查与 README：组员表、生产 `SECRET_KEY` 等（`.env.example`、`instance/` 已忽略时勾选）。

---

## 7. 前端与 AJAX 约定

- 使用 **jQuery** 发起 `$.ajax`，`contentType: 'application/json'`，重要变更带 CSRF header（如 `X-CSRFToken` 从 meta 或 cookie 读取，与 Flask 配置一致）。
- 周课表建议 **先渲染空网格**，再 `GET /api/timetable/events?week_start=...` 填充，避免整页刷新。
- 错误统一返回 JSON：`{ "error": "message", "code": "..." }`，前端 toast 或行内提示。

---

## 8. AI Planner：占位实现与接口契约（交给组员）

### 8.1 原则

- 路由与模块名保持稳定：**组员只替换占位函数内部**，不改 URL 的情况下前端无需大改。
- **禁止**把 API Key 放进前端；所有模型调用必须在服务端。

### 8.2 建议端点（已实现占位时可返回示例数据）

**`POST /api/planner/chat`**

- **Request JSON**（字段可扩展，占位实现应忽略未知字段）：

```json
{
  "message": "用户本轮输入文本",
  "conversation_id": null,
  "context": {
    "user_id": 1,
    "locale": "en",
    "upcoming_exams": [],
    "calendar_summary": null
  }
}
```

- **Response JSON**（占位阶段固定结构，便于前端解析）：

```json
{
  "ok": true,
  "stub": true,
  "reply_text": "AI Planner is not configured yet. Your teammate will connect the model here.",
  "plan_blocks": [],
  "conversation_id": null
}
```

- **`plan_blocks`** 元素建议形状（与现有 `ai_planner.html` 卡片对齐）：

```json
{
  "start_label": "7:25",
  "end_label": "8:05",
  "title": "Flask Routes",
  "duration_minutes": 40
}
```

### 8.3 组员接入清单

1. 在 `app/ai_planner/`（或约定模块）实现 `handle_chat_request(payload) -> dict`。
2. 从环境变量读取 `PLANNER_API_KEY`（名称可自定，写入 `.env.example` 说明）。
3. 将占位蓝图中的 `stub` 改为调用上述函数；失败时 `ok: false` 并返回 `error`。
4. （可选）持久化 `conversation_id` 与消息历史。

### 8.4 前端 `ai_planner.html`

- 发送按钮改为 `$.ajax` 调 `/api/planner/chat`；在 `stub: true` 时仍渲染 `reply_text`，`plan_blocks` 为空则隐藏卡片。
- 「Clear」「History」可先只做前端清空；等组员接 API 后再接会话 ID。

---

## 9. 安全与配置清单

- [ ] `SECRET_KEY`、`PLANNER_API_KEY`（将来）仅环境变量或 `instance/config.py`（不提交）。
- [ ] 登录限制：可选简单速率限制或延迟（加分项非必须）。
- [ ] 所有 POST/PATCH/DELETE 校验 CSRF 与登录态。
- [ ] 小组 join_code 防爆破：可限制尝试次数（可选）。

---

## 10. 与 Checkpoint 的对应关系（课程周次仅供参考）

| Checkpoint | 建议交付物 |
|------------|------------|
| CP2 前端 | 静态页完善 + 部分 Jinja 化 + AJAX 原型（可仍接 mock JSON） |
| CP3 后端 | Phase 1–4 核心跑通；真实 DB；AI 可为 stub |
| 截止前 | Phase 5–6、README、测试绿、演示脚本 |

---

## 11. 下一步（与仓库同步，2026-04）

1. 队友：**AI Planner** 按第 8 节替换 `stub` 实现，前端保持 `POST /api/planner/chat`。
2. 产品：对照 **`source_pages/`** 的设想与 **[`docs/SOURCE_PAGES_GAP_ANALYSIS.md`](SOURCE_PAGES_GAP_ANALYSIS.md)** 决定组页面 **Book**、考试资源表、侧栏 **Courses/Reminders** 等待办。
3. 交作业前：填 README 组员、再次跑通 **`pytest`（含 Selenium，需 Chrome）**。

---

*文档版本：与仓库同步维护；AI 契约变更时请同步本文件第 8 节并通知前端负责人。*
