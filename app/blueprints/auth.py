from flask import Blueprint, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_user, logout_user

from app.extensions import db
from app.forms import LoginForm, RegisterForm
from app.models import User

bp = Blueprint("auth", __name__)


@bp.get("/login")
def login_page():
    if current_user.is_authenticated:
        return redirect(url_for("main.timetable"))
    tab = request.args.get("tab", "login")
    return render_template(
        "auth/login_register.html",
        tab=tab if tab in ("login", "register") else "login",
        login_form=LoginForm(prefix="login"),
        register_form=RegisterForm(prefix="reg"),
    )


@bp.post("/login")
def login_submit():
    if current_user.is_authenticated:
        return redirect(url_for("main.timetable"))
    form = LoginForm(prefix="login")
    register_form = RegisterForm(prefix="reg")
    if form.validate_on_submit():
        user = User.query.filter_by(email=form.email.data.lower().strip()).first()
        if user and user.check_password(form.password.data):
            login_user(user, remember=True)
            nxt = request.args.get("next")
            if nxt and nxt.startswith("/"):
                return redirect(nxt)
            return redirect(url_for("main.timetable"))
        flash("Invalid email or password.", "danger")
    return render_template(
        "auth/login_register.html",
        tab="login",
        login_form=form,
        register_form=register_form,
    )


@bp.post("/register")
def register_submit():
    if current_user.is_authenticated:
        return redirect(url_for("main.timetable"))
    form = RegisterForm(prefix="reg")
    login_form = LoginForm(prefix="login")
    if form.validate_on_submit():
        email = form.email.data.lower().strip()
        if User.query.filter_by(email=email).first():
            flash("That email is already registered.", "warning")
            return render_template(
                "auth/login_register.html",
                tab="register",
                login_form=login_form,
                register_form=form,
            )
        user = User(
            email=email,
            full_name=form.full_name.data.strip(),
            student_id=(form.student_id.data or "").strip() or None,
        )
        user.set_password(form.password.data)
        db.session.add(user)
        db.session.commit()
        login_user(user, remember=True)
        return redirect(url_for("main.timetable"))
    return render_template(
        "auth/login_register.html",
        tab="register",
        login_form=login_form,
        register_form=form,
    )


@bp.post("/logout")
def logout():
    logout_user()
    return redirect(url_for("auth.login_page"))
