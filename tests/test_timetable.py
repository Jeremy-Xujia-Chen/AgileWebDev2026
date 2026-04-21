from datetime import datetime, timedelta


def _register(client, email: str, password: str = "password1") -> None:
    client.post(
        "/register",
        data={
            "reg-full_name": "T",
            "reg-email": email,
            "reg-password": password,
        },
    )


def test_api_events_requires_login(client):
    r = client.get("/api/timetable/events?week_start=2026-04-14")
    assert r.status_code == 401


def test_events_bad_week_start(client):
    _register(client, "badweek@student.uwa.edu.au")
    assert client.get("/api/timetable/events").status_code == 400
    assert client.get("/api/timetable/events?week_start=not-a-date").status_code == 400


def test_events_crud_week_filter(client):
    _register(client, "cal@student.uwa.edu.au")
    monday = "2026-04-14"
    empty = client.get(f"/api/timetable/events?week_start={monday}")
    assert empty.status_code == 200
    assert empty.get_json()["events"] == []

    body = {
        "title": "CITS3403 Lecture",
        "event_type": "lecture",
        "start_at": "2026-04-15T09:00:00",
        "end_at": "2026-04-15T10:00:00",
        "location": "LT 225",
        "notes": "Web dev",
    }
    c = client.post("/api/timetable/events", json=body)
    assert c.status_code == 201
    ev = c.get_json()["event"]
    assert ev["title"] == body["title"]
    assert ev["event_type"] == "lecture"
    eid = ev["id"]

    listed = client.get(f"/api/timetable/events?week_start={monday}")
    assert listed.status_code == 200
    events = listed.get_json()["events"]
    assert len(events) == 1
    assert events[0]["id"] == eid

    prev_week = client.get("/api/timetable/events?week_start=2026-04-07")
    assert prev_week.get_json()["events"] == []

    patch = client.patch(
        f"/api/timetable/events/{eid}",
        json={"title": "Updated title", "location": "Online"},
    )
    assert patch.status_code == 200
    assert patch.get_json()["event"]["title"] == "Updated title"
    assert patch.get_json()["event"]["location"] == "Online"

    dele = client.delete(f"/api/timetable/events/{eid}")
    assert dele.status_code == 200
    assert client.get(f"/api/timetable/events?week_start={monday}").get_json()["events"] == []


def test_create_event_validation(client):
    _register(client, "val@student.uwa.edu.au")
    r = client.post(
        "/api/timetable/events",
        json={
            "title": "X",
            "event_type": "lecture",
            "start_at": "2026-04-15T10:00:00",
            "end_at": "2026-04-15T09:00:00",
        },
    )
    assert r.status_code == 400

    r2 = client.post(
        "/api/timetable/events",
        json={
            "title": "X",
            "event_type": "not-a-type",
            "start_at": "2026-04-15T09:00:00",
            "end_at": "2026-04-15T10:00:00",
        },
    )
    assert r2.status_code == 400


def test_other_user_cannot_delete_event(client, app):
    _register(client, "owner@student.uwa.edu.au")
    c = client.post(
        "/api/timetable/events",
        json={
            "title": "Private",
            "event_type": "lab",
            "start_at": "2026-04-16T11:00:00",
            "end_at": "2026-04-16T13:00:00",
        },
    )
    eid = c.get_json()["event"]["id"]
    client.post("/logout", data={})

    other = app.test_client()
    _register(other, "other@student.uwa.edu.au")
    r = other.delete(f"/api/timetable/events/{eid}")
    assert r.status_code == 404
