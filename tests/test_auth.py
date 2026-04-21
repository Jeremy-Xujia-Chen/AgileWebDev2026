def test_register_login_logout_flow(client):
    r = client.post(
        "/register",
        data={
            "reg-full_name": "Test User",
            "reg-email": "tester@student.uwa.edu.au",
            "reg-password": "hunter2222",
        },
        follow_redirects=False,
    )
    assert r.status_code == 302
    assert r.headers["Location"].endswith("/timetable")

    me = client.get("/api/auth/me")
    assert me.status_code == 200
    body = me.get_json()
    assert body["authenticated"] is True
    assert body["user"]["email"] == "tester@student.uwa.edu.au"

    out = client.post("/logout", data={}, follow_redirects=False)
    assert out.status_code == 302

    me2 = client.get("/api/auth/me")
    assert me2.status_code == 401


def test_login_bad_password(client):
    client.post(
        "/register",
        data={
            "reg-full_name": "A",
            "reg-email": "a@student.uwa.edu.au",
            "reg-password": "aaaaaaaa",
        },
    )
    client.post("/logout", data={})
    r = client.post(
        "/login",
        data={
            "login-email": "a@student.uwa.edu.au",
            "login-password": "wrongpass1",
        },
    )
    assert r.status_code == 200
    assert b"Invalid email or password" in r.data


def test_timetable_requires_login(client):
    r = client.get("/timetable", follow_redirects=False)
    assert r.status_code == 302
    assert "/login" in r.headers["Location"]
