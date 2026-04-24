from __future__ import annotations

from datetime import datetime

from flask import Blueprint, jsonify, request
from flask_login import current_user, login_required

from app.extensions import db
from app.models import Course, Reminder, UserPreference

bp = Blueprint("user_data", __name__, url_prefix="/api/user")


def _parse_dt(value) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if not isinstance(value, str):
        return None
    s = value.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


@bp.get("/courses")
@login_required
def list_courses():
    rows = Course.query.filter_by(user_id=current_user.id).order_by(Course.code.asc(), Course.id.asc()).all()
    return jsonify({"courses": [c.to_dict() for c in rows]})


@bp.post("/courses")
@login_required
def create_course():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    code = (data.get("code") or "").strip()
    title = (data.get("title") or "").strip()
    if not code or not title:
        return jsonify({"error": "code and title are required."}), 400
    c = Course(user_id=current_user.id, code=code, title=title)
    db.session.add(c)
    db.session.commit()
    return jsonify({"course": c.to_dict()}), 201


@bp.delete("/courses/<int:course_id>")
@login_required
def delete_course(course_id: int):
    c = Course.query.filter_by(id=course_id, user_id=current_user.id).first()
    if c is None:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(c)
    db.session.commit()
    return jsonify({"ok": True})


@bp.get("/reminders")
@login_required
def list_reminders():
    rows = (
        Reminder.query.filter_by(user_id=current_user.id)
        .order_by(Reminder.is_done.asc(), Reminder.id.desc())
        .all()
    )
    return jsonify({"reminders": [r.to_dict() for r in rows]})


@bp.post("/reminders")
@login_required
def create_reminder():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    if not title:
        return jsonify({"error": "title is required."}), 400
    due = _parse_dt(data.get("due_at")) if data.get("due_at") not in (None, "") else None
    r = Reminder(user_id=current_user.id, title=title, due_at=due, is_done=False)
    db.session.add(r)
    db.session.commit()
    return jsonify({"reminder": r.to_dict()}), 201


@bp.patch("/reminders/<int:reminder_id>")
@login_required
def patch_reminder(reminder_id: int):
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    r = Reminder.query.filter_by(id=reminder_id, user_id=current_user.id).first()
    if r is None:
        return jsonify({"error": "Not found."}), 404
    data = request.get_json(silent=True) or {}
    if "title" in data:
        t = (data.get("title") or "").strip()
        if not t:
            return jsonify({"error": "title cannot be empty."}), 400
        r.title = t
    if "is_done" in data:
        r.is_done = bool(data.get("is_done"))
    if "due_at" in data:
        v = data.get("due_at")
        r.due_at = _parse_dt(v) if v not in (None, "") else None
    db.session.commit()
    return jsonify({"reminder": r.to_dict()})


@bp.delete("/reminders/<int:reminder_id>")
@login_required
def delete_reminder(reminder_id: int):
    r = Reminder.query.filter_by(id=reminder_id, user_id=current_user.id).first()
    if r is None:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(r)
    db.session.commit()
    return jsonify({"ok": True})


@bp.get("/preferences")
@login_required
def get_preferences():
    p = UserPreference.query.filter_by(user_id=current_user.id).first()
    if p is None:
        return jsonify({"preferences": {"timezone": "UTC", "week_starts_on": 0}})
    return jsonify({"preferences": p.to_dict()})


@bp.put("/preferences")
@login_required
def put_preferences():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    tz = (data.get("timezone") or "UTC").strip() or "UTC"
    try:
        wk = int(data.get("week_starts_on", 0))
    except (TypeError, ValueError):
        return jsonify({"error": "week_starts_on must be an integer 0–6."}), 400
    if wk < 0 or wk > 6:
        return jsonify({"error": "week_starts_on must be 0–6 (Mon=0)."}), 400
    p = UserPreference.query.filter_by(user_id=current_user.id).first()
    if p is None:
        p = UserPreference(user_id=current_user.id, timezone=tz, week_starts_on=wk)
        db.session.add(p)
    else:
        p.timezone = tz
        p.week_starts_on = wk
    db.session.commit()
    return jsonify({"preferences": p.to_dict()})
