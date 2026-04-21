"""Merged timetables and common free slots for study groups."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Any

from app.models import CalendarEvent, GroupMember, User


def monday_of(d: date) -> date:
    return d - timedelta(days=d.weekday())


def week_bounds_mon_fri(monday: date) -> tuple[datetime, datetime]:
    """Same window as timetable API: Mon 00:00 through Sat 00:00 (exclusive end)."""
    start = datetime.combine(monday, time.min)
    end = datetime.combine(monday + timedelta(days=5), time.min)
    return start, end


def user_initials(user: User) -> str:
    parts = (user.full_name or "User").split()
    a = (parts[0][0] if parts else "U").upper()
    b = (parts[1][0] if len(parts) > 1 else "").upper()
    return (a + b)[:4]


def merged_events_for_group(group_id: int, monday: date) -> list[dict[str, Any]]:
    members = GroupMember.query.filter_by(group_id=group_id).order_by(GroupMember.joined_at.asc()).all()
    if not members:
        return []
    user_ids = [m.user_id for m in members]
    order = {uid: i for i, uid in enumerate(user_ids)}
    users = {u.id: u for u in User.query.filter(User.id.in_(user_ids)).all()}
    start, end = week_bounds_mon_fri(monday)
    events = (
        CalendarEvent.query.filter(CalendarEvent.user_id.in_(user_ids))
        .filter(CalendarEvent.start_at < end)
        .filter(CalendarEvent.end_at > start)
        .order_by(CalendarEvent.start_at.asc())
        .all()
    )
    out: list[dict[str, Any]] = []
    for ev in events:
        u = users.get(ev.user_id)
        if u is None:
            continue
        d = ev.to_dict(include_user_id=True)
        d["member_name"] = u.full_name
        d["member_initials"] = user_initials(u)
        d["member_order"] = order.get(ev.user_id, 0)
        out.append(d)
    return out


def _slot_busy_for_user(user_id: int, slot_start: datetime, slot_end: datetime) -> bool:
    return (
        CalendarEvent.query.filter_by(user_id=user_id)
        .filter(CalendarEvent.start_at < slot_end)
        .filter(CalendarEvent.end_at > slot_start)
        .first()
        is not None
    )


def common_free_slots(group_id: int, monday: date, *, slot_minutes: int = 30) -> list[dict[str, Any]]:
    """Mon–Fri 08:00–20:00 local; slot is free only if no member has a calendar overlap."""
    members = GroupMember.query.filter_by(group_id=group_id).all()
    if not members:
        return []
    user_ids = [m.user_id for m in members]
    n = len(user_ids)
    slot = timedelta(minutes=slot_minutes)
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    merged: list[dict[str, Any]] = []

    for wd in range(5):
        d = monday + timedelta(days=wd)
        day_start = datetime.combine(d, time(8, 0))
        day_end = datetime.combine(d, time(20, 0))
        t = day_start
        while t + slot <= day_end:
            slot_end = t + slot
            busy_any = False
            for uid in user_ids:
                if _slot_busy_for_user(uid, t, slot_end):
                    busy_any = True
                    break
            if not busy_any:
                if merged and merged[-1]["end"] == t.isoformat(timespec="minutes"):
                    merged[-1]["end"] = slot_end.isoformat(timespec="minutes")
                else:
                    merged.append(
                        {
                            "day": day_names[wd],
                            "date": d.isoformat(),
                            "start": t.isoformat(timespec="minutes"),
                            "end": slot_end.isoformat(timespec="minutes"),
                            "member_count": n,
                        }
                    )
            t = slot_end

    return merged


def count_events_for_user_week(user_id: int, monday: date) -> int:
    start, end = week_bounds_mon_fri(monday)
    return (
        CalendarEvent.query.filter_by(user_id=user_id)
        .filter(CalendarEvent.start_at < end)
        .filter(CalendarEvent.end_at > start)
        .count()
    )
