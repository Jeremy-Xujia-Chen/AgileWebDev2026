from datetime import date


def _register(client, email: str, full_name: str = "User T", password: str = "password1") -> None:
    client.post(
        "/register",
        data={
            "reg-full_name": full_name,
            "reg-email": email,
            "reg-password": password,
        },
    )


def _logout(client) -> None:
    client.post("/logout", data={})


def _login(client, email: str, password: str = "password1") -> None:
    client.post("/login", data={"login-email": email, "login-password": password})


def test_groups_api_requires_login(client):
    assert client.get("/api/groups/mine").status_code == 401


def test_create_join_merged_leave(client):
    """Single test client: log out before switching users (Flask test clients share one session)."""
    _register(client, "gowner@student.uwa.edu.au", "Group Owner")
    c = client.post("/api/groups/", json={"name": "Team Alpha"})
    assert c.status_code == 201
    gid = c.get_json()["group"]["id"]
    code = c.get_json()["group"]["join_code"]
    assert len(code) == 8

    assert client.get("/api/groups/mine").status_code == 200
    assert len(client.get("/api/groups/mine").get_json()["groups"]) == 1

    monday = date(2026, 4, 14)
    client.post(
        "/api/timetable/events",
        json={
            "title": "Lecture A",
            "event_type": "lecture",
            "start_at": "2026-04-15T09:00:00",
            "end_at": "2026-04-15T10:00:00",
        },
    )

    _logout(client)
    _register(client, "gmember@student.uwa.edu.au", "Member One")
    j = client.post("/api/groups/join", json={"join_code": code})
    assert j.status_code == 200
    assert j.get_json()["already_member"] is False

    merged = client.get(f"/api/groups/{gid}/merged-timetable?week_start={monday.isoformat()}")
    assert merged.status_code == 200
    evs = merged.get_json()["events"]
    assert len(evs) == 1
    assert evs[0]["title"] == "Lecture A"
    assert "member_initials" in evs[0]

    assert len(client.get("/api/groups/mine").get_json()["groups"]) == 1

    assert client.post("/api/groups/join", json={"join_code": code}).get_json()["already_member"] is True

    assert client.post(f"/api/groups/{gid}/leave", json={}).status_code == 200
    assert client.get("/api/groups/mine").get_json()["groups"] == []

    _logout(client)
    _login(client, "gowner@student.uwa.edu.au")
    assert client.post(f"/api/groups/{gid}/leave", json={}).status_code == 200
    assert client.get("/api/groups/mine").get_json()["groups"] == []


def test_join_invalid_code(client):
    _register(client, "gj@student.uwa.edu.au")
    assert client.post("/api/groups/join", json={"join_code": "ZZZZZZZZ"}).status_code == 404


def test_group_tasks(client):
    _register(client, "gtask@student.uwa.edu.au", "Task Owner")
    gid = client.post("/api/groups/", json={"name": "Taskers"}).get_json()["group"]["id"]
    code = client.get(f"/api/groups/{gid}").get_json()["group"]["join_code"]

    _logout(client)
    _register(client, "gtask2@student.uwa.edu.au", "Assignee Bob")
    client.post("/api/groups/join", json={"join_code": code})
    uid_bob = client.get("/api/auth/me").get_json()["user"]["id"]

    _logout(client)
    _login(client, "gtask@student.uwa.edu.au")

    t = client.post(
        f"/api/groups/{gid}/tasks",
        json={"title": "Write README", "assignee_user_id": uid_bob, "due_date": "2026-05-01"},
    )
    assert t.status_code == 201
    tid = t.get_json()["task"]["id"]

    listed = client.get(f"/api/groups/{gid}/tasks").get_json()["tasks"]
    assert len(listed) == 1
    assert listed[0]["assignee_name"] == "Assignee Bob"

    assert (
        client.patch(
            f"/api/groups/{gid}/tasks/{tid}",
            json={"status": "done"},
        ).status_code
        == 200
    )

    assert client.delete(f"/api/groups/{gid}/tasks/{tid}").status_code == 200
    assert client.get(f"/api/groups/{gid}/tasks").get_json()["tasks"] == []


def test_non_member_no_access(client):
    _register(client, "gpriv@student.uwa.edu.au")
    gid = client.post("/api/groups/", json={"name": "Private"}).get_json()["group"]["id"]

    _logout(client)
    _register(client, "gout@student.uwa.edu.au")
    monday = date(2026, 4, 14)
    assert client.get(f"/api/groups/{gid}/merged-timetable?week_start={monday.isoformat()}").status_code == 404


def test_free_slots_all_empty_week(client):
    _register(client, "gfree@student.uwa.edu.au")
    gid = client.post("/api/groups/", json={"name": "Free Team"}).get_json()["group"]["id"]
    monday = date(2026, 6, 1)
    r = client.get(f"/api/groups/{gid}/free-slots?week_start={monday.isoformat()}")
    assert r.status_code == 200
    slots = r.get_json()["slots"]
    assert isinstance(slots, list)
    assert len(slots) >= 1


def test_group_page_renders(client):
    _register(client, "gpage@student.uwa.edu.au")
    r = client.get("/group")
    assert r.status_code == 200
    assert b"studysync_group.js" in r.data
