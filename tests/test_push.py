from __future__ import annotations

import json

import httpx
import pytest

from backend.push import PushError, assignment_push_url, push_saved_assignment


def test_assignment_push_url_empty() -> None:
    assert assignment_push_url("") is None
    assert assignment_push_url("   ") is None


def test_assignment_push_url_origin_appends_save_path() -> None:
    assert assignment_push_url("https://letters.example") == "https://letters.example/api/save"
    assert assignment_push_url("letters.example") == "https://letters.example/api/save"


def test_assignment_push_url_keeps_full_path() -> None:
    assert (
        assignment_push_url("https://letters.example/v1/assignments")
        == "https://letters.example/v1/assignments"
    )


def test_push_saved_assignment_posts_json() -> None:
    captured: dict = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["url"] = str(request.url)
        captured["body"] = json.loads(request.content)
        return httpx.Response(200, json={"ok": True})

    transport = httpx.MockTransport(handler)
    with httpx.Client(transport=transport) as client:
        result = push_saved_assignment(
            "https://letters.example/ingest",
            {"staffName": "Ada Lovelace"},
            http_client=client,
        )
    assert result["pushed"] is True
    assert captured["url"] == "https://letters.example/ingest"
    assert captured["body"]["staffName"] == "Ada Lovelace"


def test_push_saved_assignment_http_error() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, text="nope")

    transport = httpx.MockTransport(handler)
    with httpx.Client(transport=transport) as client:
        with pytest.raises(PushError):
            push_saved_assignment(
                "https://letters.example/ingest",
                {"staffName": "Ada"},
                http_client=client,
            )
