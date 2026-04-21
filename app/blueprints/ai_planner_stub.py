"""
AI Planner — stub implementation.

Teammate replaces the body of `build_stub_response` with a real model call.
Contract is documented in docs/PROJECT_IMPLEMENTATION_PLAN.md §8.
"""

from __future__ import annotations

from typing import Any

from flask import Blueprint, jsonify, request

bp = Blueprint("ai_planner", __name__)


def build_stub_response(payload: dict[str, Any] | None) -> dict[str, Any]:
    """Return the stable JSON shape the frontend should parse."""
    _ = payload  # reserved for teammate: conversation_id, context, etc.
    return {
        "ok": True,
        "stub": True,
        "reply_text": (
            "AI Planner is not connected yet. "
            "Your teammate should wire the model in app/blueprints/ai_planner_stub.py "
            "(see docs/PROJECT_IMPLEMENTATION_PLAN.md)."
        ),
        "plan_blocks": [],
        "conversation_id": None,
    }


@bp.post("/chat")
def chat():
    if not request.is_json:
        return jsonify({"ok": False, "error": "Expected application/json"}), 400
    data = request.get_json(silent=True) or {}
    return jsonify(build_stub_response(data))
