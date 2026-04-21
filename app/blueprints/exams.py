from __future__ import annotations

from datetime import datetime
from typing import Any

from flask import Blueprint, jsonify, render_template, request
from flask_login import current_user, login_required
from sqlalchemy import func, select

from app.extensions import db
from app.models import ExamSession, RevisionTopic

bp = Blueprint("exams", __name__)
api_bp = Blueprint("exams_api", __name__, url_prefix="/api/exams")


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


@bp.get("/exams")
@login_required
def list_exams():
    return render_template("exams/list.html")


@bp.get("/exams/<int:exam_id>")
@login_required
def exam_detail(exam_id: int):
    exam = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first_or_404()
    topics = exam.topics.order_by(RevisionTopic.sort_order.asc(), RevisionTopic.id.asc()).all()
    return render_template("exams/detail.html", exam=exam, topics=topics)


# --- API ---


@api_bp.get("/sessions")
@login_required
def api_list_sessions():
    rows = (
        ExamSession.query.filter_by(user_id=current_user.id)
        .order_by(ExamSession.starts_at.asc())
        .all()
    )
    return jsonify({"sessions": [r.to_dict() for r in rows]})


@api_bp.post("/sessions")
@login_required
def api_create_session():
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    if not title:
        return jsonify({"error": "title is required."}), 400
    starts_at = _parse_dt(data.get("starts_at"))
    ends_at = _parse_dt(data.get("ends_at"))
    if starts_at is None or ends_at is None:
        return jsonify({"error": "starts_at and ends_at must be valid ISO datetimes."}), 400
    if ends_at <= starts_at:
        return jsonify({"error": "ends_at must be after starts_at."}), 400
    wt = data.get("weight_percent")
    weight = None
    if wt is not None and wt != "":
        try:
            weight = float(wt)
        except (TypeError, ValueError):
            return jsonify({"error": "weight_percent must be a number."}), 400

    ex = ExamSession(
        user_id=current_user.id,
        title=title,
        course_code=(data.get("course_code") or "").strip() or None,
        starts_at=starts_at,
        ends_at=ends_at,
        location=(data.get("location") or "").strip() or None,
        weight_percent=weight,
        notes=(data.get("notes") or "").strip() or None,
    )
    db.session.add(ex)
    db.session.commit()
    return jsonify({"session": ex.to_dict()}), 201


@api_bp.get("/sessions/<int:exam_id>")
@login_required
def api_get_session(exam_id: int):
    ex = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    return jsonify({"session": ex.to_dict()})


@api_bp.patch("/sessions/<int:exam_id>")
@login_required
def api_patch_session(exam_id: int):
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    ex = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    data = request.get_json(silent=True) or {}
    if "title" in data:
        t = (data.get("title") or "").strip()
        if not t:
            return jsonify({"error": "title cannot be empty."}), 400
        ex.title = t
    if "course_code" in data:
        ex.course_code = (data.get("course_code") or "").strip() or None
    if "starts_at" in data:
        st = _parse_dt(data.get("starts_at"))
        if st is None:
            return jsonify({"error": "Invalid starts_at."}), 400
        ex.starts_at = st
    if "ends_at" in data:
        en = _parse_dt(data.get("ends_at"))
        if en is None:
            return jsonify({"error": "Invalid ends_at."}), 400
        ex.ends_at = en
    if "location" in data:
        ex.location = (data.get("location") or "").strip() or None
    if "weight_percent" in data:
        wt = data.get("weight_percent")
        if wt in (None, ""):
            ex.weight_percent = None
        else:
            try:
                ex.weight_percent = float(wt)
            except (TypeError, ValueError):
                db.session.rollback()
                return jsonify({"error": "weight_percent must be a number."}), 400
    if "notes" in data:
        ex.notes = (data.get("notes") or "").strip() or None
    if ex.ends_at <= ex.starts_at:
        db.session.rollback()
        return jsonify({"error": "ends_at must be after starts_at."}), 400
    db.session.commit()
    return jsonify({"session": ex.to_dict()})


@api_bp.delete("/sessions/<int:exam_id>")
@login_required
def api_delete_session(exam_id: int):
    ex = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(ex)
    db.session.commit()
    return jsonify({"ok": True})


@api_bp.get("/sessions/<int:exam_id>/topics")
@login_required
def api_list_topics(exam_id: int):
    ex = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    rows = ex.topics.order_by(RevisionTopic.sort_order.asc(), RevisionTopic.id.asc()).all()
    return jsonify({"topics": [t.to_dict() for t in rows]})


@api_bp.post("/sessions/<int:exam_id>/topics")
@login_required
def api_create_topic(exam_id: int):
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    ex = ExamSession.query.filter_by(id=exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    data = request.get_json(silent=True) or {}
    label = (data.get("label") or "").strip()
    if not label:
        return jsonify({"error": "label is required."}), 400
    mx = db.session.scalar(
        select(func.max(RevisionTopic.sort_order)).where(RevisionTopic.exam_id == exam_id)
    )
    sort_order = int(mx or 0) + 1
    tp = RevisionTopic(exam_id=exam_id, label=label, sort_order=sort_order)
    db.session.add(tp)
    db.session.commit()
    return jsonify({"topic": tp.to_dict()}), 201


ALLOWED_TOPIC_STATUS = frozenset({"not_started", "in_progress", "done"})


@api_bp.patch("/topics/<int:topic_id>")
@login_required
def api_patch_topic(topic_id: int):
    if not request.is_json:
        return jsonify({"error": "Expected application/json"}), 400
    tp = db.session.get(RevisionTopic, topic_id)
    if tp is None:
        return jsonify({"error": "Not found."}), 404
    ex = ExamSession.query.filter_by(id=tp.exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    data = request.get_json(silent=True) or {}
    if "label" in data:
        lab = (data.get("label") or "").strip()
        if not lab:
            return jsonify({"error": "label cannot be empty."}), 400
        tp.label = lab
    if "status" in data:
        st = (data.get("status") or "").strip()
        if st not in ALLOWED_TOPIC_STATUS:
            return jsonify({"error": "Invalid status."}), 400
        tp.status = st
    if "progress_percent" in data:
        try:
            p = int(data.get("progress_percent"))
        except (TypeError, ValueError):
            return jsonify({"error": "progress_percent must be an integer."}), 400
        tp.progress_percent = max(0, min(100, p))
    db.session.commit()
    return jsonify({"topic": tp.to_dict()})


@api_bp.delete("/topics/<int:topic_id>")
@login_required
def api_delete_topic(topic_id: int):
    tp = db.session.get(RevisionTopic, topic_id)
    if tp is None:
        return jsonify({"error": "Not found."}), 404
    ex = ExamSession.query.filter_by(id=tp.exam_id, user_id=current_user.id).first()
    if ex is None:
        return jsonify({"error": "Not found."}), 404
    db.session.delete(tp)
    db.session.commit()
    return jsonify({"ok": True})
