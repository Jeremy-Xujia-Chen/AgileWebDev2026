#!/usr/bin/env python3
"""
Quick smoke checks using Flask's test client (no browser, no running server).

Usage (from repo root):
  ./scripts/smoke_checks.py
  python3 scripts/smoke_checks.py

Exit code 0 if all checks pass, 1 otherwise.
"""
from __future__ import annotations

import os
import sys


def main() -> int:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if root not in sys.path:
        sys.path.insert(0, root)

    from app import create_app

    app = create_app("app.config.TestConfig")
    failures: list[str] = []
    client = app.test_client()

    def check(name: str, cond: bool, detail: str = "") -> None:
        if cond:
            print(f"OK  {name}" + (f" — {detail}" if detail else ""))
        else:
            print(f"FAIL {name}" + (f" — {detail}" if detail else ""))
            failures.append(name)

    r = client.get("/health")
    check("health", r.status_code == 200 and r.get_json().get("status") == "ok", str(r.get_json()))

    r = client.post("/api/planner/chat", json={"message": "hi"})
    body = r.get_json() or {}
    check(
        "ai_stub",
        r.status_code == 200 and body.get("stub") is True,
        str(body.get("reply_text", ""))[:80],
    )

    r = client.get("/api/timetable/events?week_start=2026-04-14")
    check("timetable_unauth", r.status_code == 401)

    client.post(
        "/register",
        data={
            "reg-full_name": "Smoke User",
            "reg-email": "smoke@student.uwa.edu.au",
            "reg-password": "smokepass1",
        },
    )
    r = client.post(
        "/api/timetable/events",
        json={
            "title": "Smoke lecture",
            "event_type": "lecture",
            "start_at": "2026-04-15T09:00:00",
            "end_at": "2026-04-15T10:00:00",
            "location": "X",
        },
    )
    check("timetable_create", r.status_code == 201, f"id={r.get_json().get('event', {}).get('id')}")

    r = client.get("/api/timetable/events?week_start=2026-04-14")
    evs = (r.get_json() or {}).get("events") or []
    check("timetable_list", r.status_code == 200 and len(evs) == 1, f"count={len(evs)}")

    if failures:
        print(f"\n{len(failures)} check(s) failed: {', '.join(failures)}")
        return 1
    print("\nAll smoke checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
