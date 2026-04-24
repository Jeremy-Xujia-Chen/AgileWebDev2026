# `source_pages/` 与当前实现的差异分析

`source_pages/` 下的 HTML 是**产品设想稿**（静态、自包含 CSS），用于对齐课程 UI 与信息架构。运行中的应用在 `templates/` + `static/`，由 Flask 与 AJAX 提供数据。以下**除 AI Planner 外**逐页分析：是否仍值得实现、建议做法。

**AI Planner**（`source_pages/ai_planner.html`）不展开：对话、真实模型、计划卡片的业务由队友按 `docs/PROJECT_IMPLEMENTATION_PLAN.md` 第 8 节与 `app/blueprints/ai_planner_stub.py` 接入；当前 `templates/main/ai_planner.html` 为占位即可。

---

## 1. `login.html` ↔ 认证页

| 设想稿 | 现状 | 是否还要做 |
|--------|------|------------|
| 双栏营销 + 登录/注册 Tab | `templates/auth/login_register.html` 已具备 Tab、品牌、字段 | **可选打磨**：与设想稿完全一致的左右分栏比例、背景装饰（非功能必需） |
| 第三方「Google」登录 | 未实现、按钮为 disabled | **不实现**（课程未要求）或标为未来功能即可 |

**实现建议**：保持现页；若时间允许，从 `source_pages/login.html` 抽颜色与间距到 `static/css/studysync_auth.css`，避免复制整页 HTML。

---

## 2. `timetable.html` ↔ 周课表

| 设想稿 | 现状 | 是否还要做 |
|--------|------|------------|
| 周导航、类型筛选、五列日视图、Modal | `main/timetable.html` + `studysync_timetable.js` 已实现，含类型筛选条 | **可选**：设想稿中「列表/网格」双视图（当前网格按钮为 disabled） |
| 与静态稿像素级一致 | 已高度对齐 StudySync 主题 | 按需微调 CSS |

**实现建议**：若作业强调「多视图」，再实现第二视图（同 API，不同 DOM 渲染）；否则可维持现状。

---

## 3. `exam_detail.html` ↔ 考试备考

| 设想稿 | 现状 | 是否还要做 |
|--------|------|------------|
| 英雄区、倒计时、复习主题、笔记、资源链接 | `ExamSession` + `RevisionTopic`；`exams/list.html`、`exams/detail.html` + API；考试级 `notes` 在 `ExamSession` | **部分可做**：独立 `exam_resources` 表、多条资源链接、按主题的笔记拆分（需迁移） |
| 分享只读 / share_token | 未实现 | **可选** Phase 3：生成 token、只读路由，满足「查看他人」叙事 |

**实现建议**：以当前详情页为准；先保证清单（topics）与 CRUD 稳定。资源表与 `share_token` 属加分项，按时间裁切。

---

## 4. `group.html` ↔ 我的小组

| 设想稿 | 现状 | 是否还要做 |
|--------|------|------------|
| 创建/加入、成员卡、在线状态 | `StudyGroup`、join code、成员 API；`main/group.html` 无「在线」概念 | **在线状态**：无后端支撑；若要做，可事后用简单「最近活动」或不做 |
| 合并课表、公共空闲 | `merged-timetable`、`free-slots` API + 页面已接 | 维护与性能即可 |
| 空闲槽 **Book** 落日历 | 设想稿有 Book 按钮 | **可做**：`POST` 在对应成员上创建同一段 `CalendarEvent` 或引入 `group_id`（表扩展见实现方案 §5） |

**实现建议**：优先保证合并视图与任务列表稳定；Book 为明确增值项，单独 PR。

---

## 5. 侧栏全局项（各 `source_pages` 共有）

| 入口 | 现状 | 是否还要做 |
|------|------|------------|
| Courses | 占位 `href="#"` | **Phase 5**（`courses` 表、与事件/考试关联） |
| Reminders | 占位 | 可选，提醒表或先不做 |
| Preferences | 占位 | 可选，时区/周起始日 |

---

## 6. 建议优先级（不碰 AI 核心开发）

1. **稳定与作业**：E2E（Selenium）、README 组员、`.env` 与密钥习惯。  
2. **高价值小功能**：组空闲 **Book**、考试 **resources** 或只读 **share_token**（三选一或按演示需要）。  
3. **视时间**：课表第二视图、登录页视觉 1:1、侧栏 Courses。  

本文件可随 `source_pages/` 与仓库实现同步增删，不必一次写完所有设想功能。
