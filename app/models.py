from __future__ import annotations

from flask_login import UserMixin

from app.extensions import db, login_manager


class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    student_id = db.Column(db.String(32), nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    calendar_events = db.relationship(
        "CalendarEvent",
        backref="user",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    exam_sessions = db.relationship(
        "ExamSession",
        backref="user",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    group_memberships = db.relationship(
        "GroupMember",
        backref="user",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    courses = db.relationship(
        "Course",
        backref="user",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    reminders = db.relationship(
        "Reminder",
        backref="user",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    preference = db.relationship(
        "UserPreference",
        backref="user",
        uselist=False,
        cascade="all, delete-orphan",
    )

    def set_password(self, password: str) -> None:
        from werkzeug.security import generate_password_hash

        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        from werkzeug.security import check_password_hash

        return check_password_hash(self.password_hash, password)


class CalendarEvent(db.Model):
    __tablename__ = "calendar_events"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    title = db.Column(db.String(200), nullable=False)
    event_type = db.Column(db.String(32), nullable=False, default="other", index=True)
    start_at = db.Column(db.DateTime, nullable=False, index=True)
    end_at = db.Column(db.DateTime, nullable=False, index=True)
    location = db.Column(db.String(200), nullable=True)
    notes = db.Column(db.Text, nullable=True)

    def to_dict(self, *, include_user_id: bool = False) -> dict:
        d = {
            "id": self.id,
            "title": self.title,
            "event_type": self.event_type,
            "start_at": self.start_at.isoformat(timespec="minutes"),
            "end_at": self.end_at.isoformat(timespec="minutes"),
            "location": self.location or "",
            "notes": self.notes or "",
        }
        if include_user_id:
            d["user_id"] = self.user_id
        return d


class ExamSession(db.Model):
    __tablename__ = "exam_sessions"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    title = db.Column(db.String(200), nullable=False)
    course_code = db.Column(db.String(32), nullable=True)
    starts_at = db.Column(db.DateTime, nullable=False, index=True)
    ends_at = db.Column(db.DateTime, nullable=False)
    location = db.Column(db.String(200), nullable=True)
    weight_percent = db.Column(db.Float, nullable=True)
    notes = db.Column(db.Text, nullable=True)

    topics = db.relationship(
        "RevisionTopic",
        backref="exam",
        lazy="dynamic",
        cascade="all, delete-orphan",
        order_by="RevisionTopic.sort_order",
    )

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "title": self.title,
            "course_code": self.course_code or "",
            "starts_at": self.starts_at.isoformat(timespec="minutes"),
            "ends_at": self.ends_at.isoformat(timespec="minutes"),
            "location": self.location or "",
            "weight_percent": self.weight_percent,
            "notes": self.notes or "",
        }


class RevisionTopic(db.Model):
    __tablename__ = "revision_topics"

    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey("exam_sessions.id"), nullable=False, index=True)
    label = db.Column(db.String(300), nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    status = db.Column(db.String(20), nullable=False, default="not_started")
    progress_percent = db.Column(db.Integer, nullable=False, default=0)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "exam_id": self.exam_id,
            "label": self.label,
            "sort_order": self.sort_order,
            "status": self.status,
            "progress_percent": self.progress_percent,
        }


class Course(db.Model):
    __tablename__ = "courses"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    code = db.Column(db.String(32), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    def to_dict(self) -> dict:
        return {"id": self.id, "code": self.code, "title": self.title}


class Reminder(db.Model):
    __tablename__ = "reminders"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    title = db.Column(db.String(200), nullable=False)
    is_done = db.Column(db.Boolean, nullable=False, default=False)
    due_at = db.Column(db.DateTime, nullable=True, index=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "title": self.title,
            "is_done": self.is_done,
            "due_at": self.due_at.isoformat(timespec="minutes") if self.due_at else None,
        }


class UserPreference(db.Model):
    __tablename__ = "user_preferences"

    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), primary_key=True)
    timezone = db.Column(db.String(64), nullable=False, default="UTC")
    week_starts_on = db.Column(db.Integer, nullable=False, default=0)

    def to_dict(self) -> dict:
        return {"timezone": self.timezone, "week_starts_on": self.week_starts_on}


class StudyGroup(db.Model):
    __tablename__ = "study_groups"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    join_code = db.Column(db.String(16), unique=True, nullable=False, index=True)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    members = db.relationship(
        "GroupMember",
        back_populates="group",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )
    tasks = db.relationship(
        "GroupTask",
        back_populates="group",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )

    def to_dict(self, *, include_join_code: bool = True) -> dict:
        d = {
            "id": self.id,
            "name": self.name,
            "created_by_user_id": self.created_by_user_id,
            "member_count": self.members.count(),
        }
        if include_join_code:
            d["join_code"] = self.join_code
        return d


class GroupMember(db.Model):
    __tablename__ = "group_members"
    __table_args__ = (db.UniqueConstraint("group_id", "user_id", name="uq_group_member_user"),)

    id = db.Column(db.Integer, primary_key=True)
    group_id = db.Column(db.Integer, db.ForeignKey("study_groups.id"), nullable=False, index=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    role = db.Column(db.String(20), nullable=False, default="member")
    joined_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    group = db.relationship("StudyGroup", back_populates="members")


class GroupTask(db.Model):
    __tablename__ = "group_tasks"

    id = db.Column(db.Integer, primary_key=True)
    group_id = db.Column(db.Integer, db.ForeignKey("study_groups.id"), nullable=False, index=True)
    title = db.Column(db.String(300), nullable=False)
    assignee_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True, index=True)
    due_date = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(20), nullable=False, default="not_started")
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    created_at = db.Column(db.DateTime, server_default=db.func.now(), nullable=False)

    group = db.relationship("StudyGroup", back_populates="tasks")

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "group_id": self.group_id,
            "title": self.title,
            "assignee_user_id": self.assignee_user_id,
            "due_date": self.due_date.isoformat() if self.due_date else None,
            "status": self.status,
            "created_by_user_id": self.created_by_user_id,
        }


@login_manager.user_loader
def load_user(user_id: str) -> User | None:
    if not user_id:
        return None
    return db.session.get(User, int(user_id))
