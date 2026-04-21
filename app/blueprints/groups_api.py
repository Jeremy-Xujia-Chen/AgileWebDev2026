from __future__ import annotations

import secrets
import string
from datetime import date, datetime

from flask import Blueprint, jsonify, request
from flask_login import current_user, login_required

from app.extensions import db
from app.models import GroupMember, GroupTask, StudyGroup, User
from app.services.group_schedule import (
    common_free_slots,
    count_events_for_user_week,
    merged_events_for_group,
    monday_of,
    user_initials,
)

bp = Blueprint("groups_api", __name__, url_prefix="/api/groups")

_ALPH = string.ascii_uppercase + string.digits
TASK_STATUSES = frozenset({"not_started", "in_progress", "done"})


def _new_join_code() -> str:
    for _ in range(64):
        code = "".join(secrets.choice(_ALPH) for _ in range(8))
        if db.session.query(StudyGroup.id).filter_by(join_code=code).first() is None:
            return code
    raise RuntimeError("join code allocation failed")


def _parse_week_start(raw: str | None) -> date | None:
    if not raw:
        return None
    try:
        d = date.fromisoformat(raw)
    except ValueError:
        return None
    return monday_of(d)


def _member_of(group_id: int) -> GroupMember | None:
    return GroupMember.query.filter_by(group_id=group_id, user_id=current_user.id).first()


def _group_if_member(group_id: int) -> StudyGroup | None:
    g = db.session.get(StudyGroup, group_id)
    if g is None:
        return None
    if _member_of(group_id) is None:
        return None
    return g


def _member_user_ids(group_id: int) -> set[int]:
    return {m.user_id for m in GroupMember.query.filter_by(group_id=group_id).all()}


@bp.get("/mine")
@login_required
def list_my_groups():
    rows = (
        db.session.query(StudyGroup)
        .join(GroupMember, GroupMember.group_id == StudyGroup.id)
        .filter(GroupMember.user_id == current_user.id)
        .order_by(StudyGroup.name.asc())
        .all()
    )
    out = []
    for g in rows:
        gm = GroupMember.query.filter_by(group_id=g.id, user_id=current_user.id).first()
        d = g.to_dict(include_join_code=True)
        d["role"] = gm.role if gm else "member"
        out.append(d)
    return jsonify({"groups": out})


@bp.post("/")
@login_required
def create_group():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"error": "name is required."}), 400
    g = StudyGroup(name=name, join_code=_new_join_code(), created_by_user_id=current_user.id)
    db.session.add(g)
    db.session.flush()
    db.session.add(GroupMember(group_id=g.id, user_id=current_user.id, role="owner"))
    db.session.commit()
    return jsonify({"group": g.to_dict(include_join_code=True)}), 201


@bp.post("/join")
@login_required
def join_group():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    raw = (data.get("join_code") or "").strip().upper()
    if not raw:
        return jsonify({"error": "join_code is required."}), 400
    g = StudyGroup.query.filter_by(join_code=raw).first()
    if g is None:
        return jsonify({"error": "Invalid join code."}), 404
    existing = GroupMember.query.filter_by(group_id=g.id, user_id=current_user.id).first()
    if existing:
        return jsonify({"group": g.to_dict(include_join_code=True), "already_member": True})
    db.session.add(GroupMember(group_id=g.id, user_id=current_user.id, role="member"))
    db.session.commit()
    return jsonify({"group": g.to_dict(include_join_code=True), "already_member": False})


@bp.get("/<int:group_id>")
@login_required
def get_group(group_id: int):
    g = _group_if_member(group_id)
    if g is None:
        return jsonify({"error": "Not found."}), 404
    ws = _parse_week_start(request.args.get("week_start")) or monday_of(date.today())
    members_out = []
    for m in GroupMember.query.filter_by(group_id=group_id).order_by(GroupMember.joined_at.asc()).all():
        u = db.session.get(User, m.user_id)
        if u is None:
            continue
        members_out.append(
            {
                "user_id": u.id,
                "full_name": u.full_name,
                "student_id": u.student_id or "",
                "initials": user_initials(u),
                "role": m.role,
                "is_you": u.id == current_user.id,
                "events_this_week": count_events_for_user_week(u.id, ws),
            }
        )
    return jsonify(
        {
            "group": g.to_dict(include_join_code=True),
            "members": members_out,
            "metrics_week_start": ws.isoformat(),
        }
    )


@bp.post("/<int:group_id>/leave")
@login_required
def leave_group(group_id: int):
    g = db.session.get(StudyGroup, group_id)
    if g is None:
        return jsonify({"error": "Not found."}), 404
    gm = GroupMember.query.filter_by(group_id=group_id, user_id=current_user.id).first()
    if gm is None:
        return jsonify({"error": "Not a member."}), 404

    others = GroupMember.query.filter(
        GroupMember.group_id == group_id,
        GroupMember.user_id != current_user.id,
    ).all()

    if not others:
        db.session.delete(g)
        db.session.commit()
        return jsonify({"ok": True, "left": True, "group_deleted": True})

    if gm.role == "owner":
        nxt = min(others, key=lambda x: (x.joined_at, x.id))
        nxt.role = "owner"
        g.created_by_user_id = nxt.user_id

    db.session.delete(gm)
    db.session.commit()
    return jsonify({"ok": True, "left": True, "group_deleted": False})


@bp.get("/<int:group_id>/merged-timetable")
@login_required
def merged_timetable(group_id: int):
    if _group_if_member(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    ws = _parse_week_start(request.args.get("week_start"))
    if ws is None:
        return jsonify({"error": "Invalid or missing week_start (YYYY-MM-DD)."}), 400
    events = merged_events_for_group(group_id, ws)
    return jsonify({"week_start": ws.isoformat(), "events": events})


@bp.get("/<int:group_id>/free-slots")
@login_required
def free_slots(group_id: int):
    if _group_if_member(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    ws = _parse_week_start(request.args.get("week_start"))
    if ws is None:
        return jsonify({"error": "Invalid or missing week_start (YYYY-MM-DD)."}), 400
    slots = common_free_slots(group_id, ws)
    return jsonify({"week_start": ws.isoformat(), "slots": slots})


@bp.get("/<int:group_id>/tasks")
@login_required
def list_tasks(group_id: int):
    if _group_if_member(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    tasks = GroupTask.query.filter_by(group_id=group_id).order_by(GroupTask.due_date.asc(), GroupTask.id.asc()).all()
    ids = {t.assignee_user_id for t in tasks if t.assignee_user_id}
    assignees = {u.id: u for u in User.query.filter(User.id.in_(ids)).all()} if ids else {}
    out = []
    for t in tasks:
        d = t.to_dict()
        aid = t.assignee_user_id
        if aid and aid in assignees:
            u = assignees[aid]
            d["assignee_name"] = u.full_name
            d["assignee_initials"] = user_initials(u)
        else:
            d["assignee_name"] = None
            d["assignee_initials"] = None
        out.append(d)
    return jsonify({"tasks": out})


@bp.post("/<int:group_id>/tasks")
@login_required
def create_task(group_id: int):
    if _group_if_member(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    if not title:
        return jsonify({"error": "title is required."}), 400
    allowed = _member_user_ids(group_id)
    assignee = data.get("assignee_user_id")
    if assignee is not None and assignee != "":
        try:
            assignee = int(assignee)
        except (TypeError, ValueError):
            return jsonify({"error": "assignee_user_id must be an integer or null."}), 400
        if assignee not in allowed:
            return jsonify({"error": "assignee must be a member of this group."}), 400
    else:
        assignee = None

    due_raw = data.get("due_date")
    due: date | None = None
    if due_raw:
        try:
            due = date.fromisoformat(str(due_raw).strip()[:10])
        except ValueError:
            return jsonify({"error": "due_date must be YYYY-MM-DD."}), 400

    st = (data.get("status") or "not_started").strip()
    if st not in TASK_STATUSES:
        return jsonify({"error": "Invalid status."}), 400

    t = GroupTask(
        group_id=group_id,
        title=title,
        assignee_user_id=assignee,
        due_date=due,
        status=st,
        created_by_user_id=current_user.id,
    )
    db.session.add(t)
    db.session.commit()
    return jsonify({"task": t.to_dict()}), 201


@bp.patch("/<int:group_id>/tasks/<int:task_id>")
@login_required
def patch_task(group_id: int, task_id: int):
    if _group_if_member(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    t = GroupTask.query.filter_by(id=task_id, group_id=group_id).first()
    if t is None:
        return jsonify({"error": "Not found."}), 404
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    allowed = _member_user_ids(group_id)

    if "title" in data:
        tt = (data.get("title") or "").strip()
        if not tt:
            return jsonify({"error": "title cannot be empty."}), 400
        t.title = tt
    if "assignee_user_id" in data:
        av = data.get("assignee_user_id")
        if av in (None, ""):
            t.assignee_user_id = None
        else:
            try:
                av = int(av)
            except (TypeError, ValueError):
                return jsonify({"error": "assignee_user_id must be an integer or null."}), 400
            if av not in allowed:
                return jsonify({"error": "assignee must be a member of this group."}), 400
            t.assignee_user_id = av
    if "due_date" in data:
        dr = data.get("due_date")
        if dr in (None, ""):
            t.due_date = None
        else:
            try:
                t.due_date = date.fromisoformat(str(dr).strip()[:10])
            except ValueError:
                return jsonify({"error": "due_date must be YYYY-MM-DD."}), 400
    if "status" in data:
        st = (data.get("status") or "").strip()
        if st not in TASK_STATUSES:
            return jsonify({"error": "Invalid status."}), 400
        t.status = st

    db.session.commit()
    return jsonify({"task": t.to_dict()})


@bp.delete("/<int:group_id>/tasks/<int:task_id>")
@login_required
def delete_task(group_id: int, task_id: int):
    if _member_of(group_id) is None:
        return jsonify({"error": "Not found."}), 404
    t = GroupTask.query.filter_by(id=task_id, group_id=group_id).first()
    if t is None:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(t)
    db.session.commit()
    return jsonify({"ok": True})
