import os
from pathlib import Path

from flask import Flask, jsonify, redirect, request, url_for

from app.config import Config
from app.extensions import csrf, db, login_manager

_ROOT = Path(__file__).resolve().parent.parent


def create_app(config_object=Config) -> Flask:
    app = Flask(
        __name__,
        instance_relative_config=True,
        template_folder=str(_ROOT / "templates"),
        static_folder=str(_ROOT / "static"),
    )
    if isinstance(config_object, str):
        app.config.from_object(config_object)
    else:
        app.config.from_object(config_object)

    os.makedirs(app.instance_path, exist_ok=True)

    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    login_manager.login_view = "auth.login_page"

    @login_manager.unauthorized_handler
    def _unauthorized():
        if request.path.startswith("/api/"):
            return jsonify({"error": "Unauthorized"}), 401
        return redirect(url_for("auth.login_page", next=request.url))

    from app import models  # noqa: F401 — registers user_loader

    from app.blueprints.ai_planner_stub import bp as ai_planner_bp
    from app.blueprints.auth import bp as auth_bp
    from app.blueprints.exams import api_bp as exams_api_bp
    from app.blueprints.exams import bp as exams_bp
    from app.blueprints.groups_api import bp as groups_api_bp
    from app.blueprints.main import bp as main_bp
    from app.blueprints.timetable_api import bp as timetable_api_bp
    from app.blueprints.user_data import bp as user_data_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(exams_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(timetable_api_bp, url_prefix="/api/timetable")
    app.register_blueprint(exams_api_bp)
    app.register_blueprint(groups_api_bp)
    app.register_blueprint(ai_planner_bp, url_prefix="/api/planner")
    app.register_blueprint(user_data_bp)

    with app.app_context():
        db.create_all()

    return app
