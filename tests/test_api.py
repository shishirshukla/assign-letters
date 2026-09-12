from __future__ import annotations

from pathlib import Path

from fastapi.testclient import TestClient

from backend.config import Settings
from backend.main import create_app, filter_staff, inject_manifest_urls


def make_client(tmp_path: Path) -> TestClient:
    staff = tmp_path / "staff.json"
    staff.write_text(
        '[{"StaffName":"Ada Lovelace","Department":"ada@contoso.com"},'
        '{"StaffName":"Alan Turing","Department":"alan@contoso.com"}]',
        encoding="utf-8",
    )
    settings = Settings(
        log_path=tmp_path / "assignletters.log",
        staff_path=staff,
        public_base_url="https://assignletters.example",
        header_name="X-AssignLetters-Id",
        missing_header_message="Header missing — cannot assign.",
        cors_allow_origins="*",
    )
    return TestClient(create_app(settings))


def test_health(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    response = client.get("/health")
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["headerName"] == "X-AssignLetters-Id"


def test_config_from_env(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    body = client.get("/api/config").json()
    assert body["headerName"] == "X-AssignLetters-Id"
    assert "cannot assign" in body["missingHeaderMessage"]
    assert body["apiUrl"] == ""


def test_staff_filter(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    all_rows = client.get("/api/staff").json()
    assert all_rows["count"] == 2
    ada = client.get("/api/staff", params={"email": "ada@contoso.com"}).json()
    assert ada["count"] == 1
    assert ada["staff"][0]["StaffName"] == "Ada Lovelace"


def test_filter_contains_email() -> None:
    rows = [{"StaffName": "Pat", "Department": "Finance <pat@bank.test>"}]
    assert filter_staff(rows, "pat@bank.test") == rows
    assert filter_staff(rows, "other@bank.test") == []


def test_save_and_logs(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    response = client.post(
        "/api/save",
        json={
            "staffName": "Ada Lovelace",
            "department": "ada@contoso.com",
            "deadlineDate": "2026-09-30",
            "internetHeaderName": "X-AssignLetters-Id",
            "internetHeaderValue": "AL-1001",
            "userEmail": "ada@contoso.com",
            "subject": "Pension letter",
        },
        headers={"Origin": "https://outlook.office.com"},
    )
    assert response.status_code == 200
    assert response.json()["status"] == "Success"
    assert response.headers.get("access-control-allow-origin") == "*"

    text = client.get("/api/logs").text
    assert "assignment_saved" in text
    assert "AL-1001" in text
    assert "Ada Lovelace" in text

    json_logs = client.get("/api/logs", params={"format": "json"}).json()
    assert json_logs["count"] >= 1


def test_cors_preflight(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    response = client.options(
        "/api/save",
        headers={
            "Origin": "https://outlook.office.com",
            "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "content-type",
        },
    )
    assert response.status_code in (200, 204)
    assert response.headers.get("access-control-allow-origin") == "*"


def test_manifest_is_read_mode(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    response = client.get("/manifest.xml")
    assert response.status_code == 200
    xml = response.text
    assert "AssignLetters" in xml
    assert "MessageReadCommandSurface" in xml
    assert "FormType=\"Read\"" in xml
    assert "ItemEdit" not in xml
    assert "MessageComposeCommandSurface" not in xml
    assert "https://assignletters.example/taskpane.html" in xml


def test_inject_manifest_urls() -> None:
    xml = inject_manifest_urls(
        "<AppDomains>\n    <AppDomain>localhost</AppDomain>\n  </AppDomains>"
        "https://localhost:8000/taskpane.html",
        "https://tunnel.example",
    )
    assert "https://tunnel.example/taskpane.html" in xml
    assert "<AppDomain>tunnel.example</AppDomain>" in xml


def test_taskpane_assets(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    html = client.get("/taskpane.html")
    assert html.status_code == 200
    assert "office.js" in html.text
    js = client.get("/taskpane.js")
    assert js.status_code == 200
    assert "getAllInternetHeadersAsync" in js.text
