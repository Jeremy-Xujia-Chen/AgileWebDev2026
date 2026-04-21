from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Any

from flask import Blueprint, jsonify, request
from flask_login import current_user, login_required

from app.extensions import db
from app.models import CalendarEvent

bp = Blueprint("timetable_api", __name__)

ALLOWED_TYPES = frozenset(
    {"lecture", "lab", "tutorial", "exam", "assignment", "workshop", "other"}
)


def _monday_of(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _week_bounds(monday: date) -> tuple[datetime, datetime]:
    start = datetime.combine(monday, time.min)
    end = datetime.combine(monday + timedelta(days=5), time.min)
    return start, end


def _parse_week_start(raw: str | None) -> date | None:
    if not raw:
        return None
    try:
        d = date.fromisoformat(raw)
    except ValueError:
        return None
    return _monday_of(d)


def _parse_dt(value: Any) -> datetime | None:
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


@bp.get("/events")
@login_required
def list_events():
    ws = _parse_week_start(request.args.get("week_start"))
    if ws is None:
        return jsonify({"error": "Invalid or missing week_start (YYYY-MM-DD)."}), 400
    start, end = _week_bounds(ws)
    q = (
        CalendarEvent.query.filter_by(user_id=current_user.id)
        .filter(CalendarEvent.start_at < end)
        .filter(CalendarEvent.end_at > start)
        .order_by(CalendarEvent.start_at.asc())
    )
    return jsonify(
        {
            "week_start": ws.isoformat(),
            "events": [e.to_dict() for e in q.all()],
        }
    )


@bp.post("/events")
@login_required
def create_event():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    event_type = (data.get("event_type") or "other").strip().lower()
    if not title:
        return jsonify({"error": "title is required."}), 400
    if event_type not in ALLOWED_TYPES:
        return jsonify({"error": f"event_type must be one of: {sorted(ALLOWED_TYPES)}."}), 400
    start_at = _parse_dt(data.get("start_at"))
    end_at = _parse_dt(data.get("end_at"))
    if start_at is None or end_at is None:
        return jsonify({"error": "start_at and end_at must be valid ISO datetimes."}), 400
    if end_at <= start_at:
        return jsonify({"error": "end_at must be after start_at."}), 400

    ev = CalendarEvent(
        user_id=current_user.id,
        title=title,
        event_type=event_type,
        start_at=start_at,
        end_at=end_at,
        location=(data.get("location") or "").strip() or None,
        notes=(data.get("notes") or "").strip() or None,
    )
    db.session.add(ev)
    db.session.commit()
    return jsonify({"event": ev.to_dict()}), 201


@bp.patch("/events/<int:event_id>")
@login_required
def update_event(event_id: int):
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    ev = db.session.get(CalendarEvent, event_id)
    if ev is None or ev.user_id != current_user.id:
        return jsonify({"error": "Not found."}), 404
    data = request.get_json(silent=True) or {}

    if "title" in data:
        t = (data.get("title") or "").strip()
        if not t:
            return jsonify({"error": "title cannot be empty."}), 400
        ev.title = t
    if "event_type" in data:
        et = (data.get("event_type") or "").strip().lower()
        if et not in ALLOWED_TYPES:
            return jsonify({"error": f"event_type must be one of: {sorted(ALLOWED_TYPES)}."}), 400
        ev.event_type = et
    if "start_at" in data:
        st = _parse_dt(data.get("start_at"))
        if st is None:
            return jsonify({"error": "Invalid start_at."}), 400
        ev.start_at = st
    if "end_at" in data:
        en = _parse_dt(data.get("end_at"))
        if en is None:
            return jsonify({"error": "Invalid end_at."}), 400
        ev.end_at = en
    if "location" in data:
        ev.location = (data.get("location") or "").strip() or None
    if "notes" in data:
        ev.notes = (data.get("notes") or "").strip() or None

    if ev.end_at <= ev.start_at:
        db.session.rollback()
        return jsonify({"error": "end_at must be after start_at."}), 400

    db.session.commit()
    return jsonify({"event": ev.to_dict()})


@bp.delete("/events/<int:event_id>")
@login_required
def delete_event(event_id: int):
    ev = db.session.get(CalendarEvent, event_id)
    if ev is None or ev.user_id != current_user.id:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(ev)
    db.session.commit()
    return jsonify({"ok": True})
