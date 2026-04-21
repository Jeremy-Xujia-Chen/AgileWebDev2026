def test_health_ok(client):
    r = client.get("/health")
    assert r.status_code == 200
    data = r.get_json()
    assert data["status"] == "ok"
    assert data["app"] == "studysync"


def test_ai_planner_chat_stub(client):
    r = client.post("/api/planner/chat", json={"message": "plan my afternoon"})
    assert r.status_code == 200
    data = r.get_json()
    assert data["ok"] is True
    assert data["stub"] is True
    assert "reply_text" in data
    assert data["plan_blocks"] == []
    assert data["conversation_id"] is None


def test_ai_planner_chat_requires_json(client):
    r = client.post("/api/planner/chat", data="not-json")
    assert r.status_code == 400


def _register(client, email: str) -> None:
    client.post(
        "/register",
        data={
            "reg-full_name": "Planner User",
            "reg-email": email,
            "reg-password": "password1",
        },
    )


def test_ai_planner_page_renders(client):
    _register(client, "planner@student.uwa.edu.au")
    r = client.get("/ai-planner")
    assert r.status_code == 200
    assert b"StudySync AI" in r.data
    assert b"studysync_ai.js" in r.data
