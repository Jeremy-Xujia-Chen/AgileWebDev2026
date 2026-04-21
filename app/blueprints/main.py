from flask import Blueprint, jsonify, redirect, render_template, url_for
from flask_login import current_user, login_required

bp = Blueprint("main", __name__)


@bp.get("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("main.timetable"))
    return redirect(url_for("auth.login_page"))


@bp.get("/timetable")
@login_required
def timetable():
    return render_template("main/timetable.html")


@bp.get("/ai-planner")
@login_required
def ai_planner_page():
    return render_template("main/ai_planner.html")


@bp.get("/group")
@login_required
def group_page():
    return render_template("main/group.html")


@bp.get("/api/auth/me")
def api_auth_me():
    """Lightweight JSON for AJAX clients (GET is CSRF-exempt)."""
    if not current_user.is_authenticated:
        return jsonify({"authenticated": False}), 401
    return jsonify(
        {
            "authenticated": True,
            "user": {
                "id": current_user.id,
                "email": current_user.email,
                "full_name": current_user.full_name,
                "student_id": current_user.student_id,
            },
        }
    )


@bp.get("/health")
def health():
    return jsonify({"status": "ok", "app": "studysync"})
