from __future__ import annotations

import pytest

from backend.push import PushUrlError, assignment_push_url


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


def test_assignment_push_url_rejects_bad_scheme() -> None:
    with pytest.raises(PushUrlError):
        assignment_push_url("ftp://letters.example/save")
