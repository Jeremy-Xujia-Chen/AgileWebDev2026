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


def test_exams_api_requires_login(client):
    assert client.get("/api/exams/sessions").status_code == 401


def test_exams_sessions_crud(client):
    _register(client, "exams1@student.uwa.edu.au")
    t0 = datetime(2026, 5, 10, 9, 0, 0)
    t1 = datetime(2026, 5, 10, 11, 0, 0)
    c = client.post(
        "/api/exams/sessions",
        json={
            "title": "Final CITS3403",
            "course_code": "CITS3403",
            "starts_at": t0.isoformat(timespec="minutes"),
            "ends_at": t1.isoformat(timespec="minutes"),
            "location": "Arts G52",
            "weight_percent": 40,
            "notes": "Bring student card",
        },
    )
    assert c.status_code == 201
    sid = c.get_json()["session"]["id"]

    listed = client.get("/api/exams/sessions")
    assert listed.status_code == 200
    assert len(listed.get_json()["sessions"]) == 1

    g = client.get(f"/api/exams/sessions/{sid}")
    assert g.status_code == 200
    assert g.get_json()["session"]["title"] == "Final CITS3403"

    p = client.patch(
        f"/api/exams/sessions/{sid}",
        json={"title": "Updated title", "notes": ""},
    )
    assert p.status_code == 200
    assert p.get_json()["session"]["title"] == "Updated title"

    tp = client.post(
        f"/api/exams/sessions/{sid}/topics",
        json={"label": "REST APIs"},
    )
    assert tp.status_code == 201
    tid = tp.get_json()["topic"]["id"]

    topics = client.get(f"/api/exams/sessions/{sid}/topics")
    assert topics.status_code == 200
    assert len(topics.get_json()["topics"]) == 1

    patch_t = client.patch(
        f"/api/exams/topics/{tid}",
        json={"status": "done", "progress_percent": 100},
    )
    assert patch_t.status_code == 200
    assert patch_t.get_json()["topic"]["status"] == "done"

    assert client.delete(f"/api/exams/topics/{tid}").status_code == 200
    assert client.get(f"/api/exams/sessions/{sid}/topics").get_json()["topics"] == []

    assert client.delete(f"/api/exams/sessions/{sid}").status_code == 200
    assert client.get("/api/exams/sessions").get_json()["sessions"] == []


def test_create_session_validation(client):
    _register(client, "exams2@student.uwa.edu.au")
    r = client.post(
        "/api/exams/sessions",
        json={
            "title": "X",
            "starts_at": "2026-06-01T14:00:00",
            "ends_at": "2026-06-01T12:00:00",
        },
    )
    assert r.status_code == 400

    r2 = client.post(
        "/api/exams/sessions",
        json={"title": "", "starts_at": "2026-06-01T12:00:00", "ends_at": "2026-06-01T14:00:00"},
    )
    assert r2.status_code == 400


def test_other_user_cannot_access_exam(client, app):
    _register(client, "owner2@student.uwa.edu.au")
    t0 = datetime(2026, 7, 1, 10, 0, 0)
    t1 = t0 + timedelta(hours=2)
    sid = client.post(
        "/api/exams/sessions",
        json={"title": "Mine", "starts_at": t0.isoformat(), "ends_at": t1.isoformat()},
    ).get_json()["session"]["id"]

    client.post("/logout", data={})
    other = app.test_client()
    _register(other, "other2@student.uwa.edu.au")
    assert other.get(f"/api/exams/sessions/{sid}").status_code == 404
    assert other.delete(f"/api/exams/sessions/{sid}").status_code == 404


def test_exams_pages_require_login(client):
    assert client.get("/exams").status_code == 302
    assert client.get("/exams/1").status_code == 302


def test_exams_pages_ok_when_logged_in(client):
    _register(client, "exams3@student.uwa.edu.au")
    assert client.get("/exams").status_code == 200
    t0 = datetime(2026, 8, 1, 9, 0, 0)
    t1 = t0 + timedelta(hours=3)
    sid = client.post(
        "/api/exams/sessions",
        json={"title": "Mid", "starts_at": t0.isoformat(), "ends_at": t1.isoformat()},
    ).get_json()["session"]["id"]
    r = client.get(f"/exams/{sid}")
    assert r.status_code == 200
    assert b"Mid" in r.data
