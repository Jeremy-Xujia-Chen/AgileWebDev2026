def _reg(client, email: str) -> None:
    client.post(
        "/register",
        data={"reg-full_name": "T", "reg-email": email, "reg-password": "password1"},
    )


def test_user_api_auth(client):
    assert client.get("/api/user/courses").status_code == 401


def test_courses_reminders_preferences(client):
    _reg(client, "u1@student.uwa.edu.au")
    assert client.get("/api/user/courses").get_json()["courses"] == []
    a = client.post("/api/user/courses", json={"code": "CITS3403", "title": "Web"})
    assert a.status_code == 201
    cid = a.get_json()["course"]["id"]
    assert client.delete(f"/api/user/courses/{cid}").status_code == 200

    r = client.post("/api/user/reminders", json={"title": "X", "due_at": "2026-08-01T12:00:00"})
    rid = r.get_json()["reminder"]["id"]
    assert client.patch(f"/api/user/reminders/{rid}", json={"is_done": True}).status_code == 200
    assert client.delete(f"/api/user/reminders/{rid}").status_code == 200

    assert client.get("/api/user/preferences").get_json()["preferences"]["timezone"] == "UTC"
    assert client.put(
        "/api/user/preferences", json={"timezone": "Australia/Perth", "week_starts_on": 0}
    ).status_code == 200


def test_user_pages_200(client):
    _reg(client, "u2@student.uwa.edu.au")
    for path in ("/courses", "/reminders", "/preferences"):
        assert client.get(path).status_code == 200
