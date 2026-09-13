from __future__ import annotations

from pathlib import Path

from fastapi.testclient import TestClient

from backend.config import Settings
from backend.main import create_app, filter_staff
from backend.manifest import ManifestUrlError, generate_manifest, normalize_public_origin


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
    assert "<AppDomain>https://assignletters.example</AppDomain>" in xml
    assert "<AppDomain>localhost</AppDomain>" not in xml
    assert "<AppDomain>assignletters.example</AppDomain>" not in xml


def test_generate_manifest_from_placeholder() -> None:
    xml = generate_manifest(
        "<OfficeApp>"
        "<AppDomains>\n    <AppDomain>{{PUBLIC_BASE_URL}}</AppDomain>\n  </AppDomains>"
        '<SourceLocation DefaultValue="{{PUBLIC_BASE_URL}}/taskpane.html" />'
        "</OfficeApp>",
        "https://tunnel.example",
    )
    assert "https://tunnel.example/taskpane.html" in xml
    assert "<AppDomain>https://tunnel.example</AppDomain>" in xml
    assert "{{PUBLIC_BASE_URL}}" not in xml


def test_normalize_origin_adds_https_and_strips_path() -> None:
    assert normalize_public_origin("tunnel.example/foo") == "https://tunnel.example"
    assert normalize_public_origin("https://tunnel.example:8443/") == "https://tunnel.example:8443"


def test_generate_rejects_empty_base() -> None:
    try:
        generate_manifest("<OfficeApp/>", "")
    except ManifestUrlError:
        return
    raise AssertionError("expected ManifestUrlError")


def test_manifest_reads_public_base_url_from_env(tmp_path: Path, monkeypatch) -> None:
    staff = tmp_path / "staff.json"
    staff.write_text("[]", encoding="utf-8")
    monkeypatch.setenv("ASSIGNLETTERS_PUBLIC_BASE_URL", "https://from-env.example")
    monkeypatch.setenv("ASSIGNLETTERS_STAFF_PATH", str(staff))
    monkeypatch.setenv("ASSIGNLETTERS_LOG_PATH", str(tmp_path / "assignletters.log"))
    client = TestClient(create_app())
    xml = client.get("/manifest.xml").text
    assert "https://from-env.example/taskpane.html" in xml
    assert "<AppDomain>https://from-env.example</AppDomain>" in xml
    assert "{{PUBLIC_BASE_URL}}" not in xml


def test_manifest_env_url_wins_over_request_host(tmp_path: Path) -> None:
    staff = tmp_path / "staff.json"
    staff.write_text("[]", encoding="utf-8")
    settings = Settings(
        log_path=tmp_path / "assignletters.log",
        staff_path=staff,
        public_base_url="https://from-env.example",
    )
    client = TestClient(create_app(settings))
    xml = client.get(
        "/manifest.xml",
        headers={
            "Host": "other-host.example",
            "X-Forwarded-Host": "other-host.example",
            "X-Forwarded-Proto": "https",
        },
    ).text
    assert "https://from-env.example/taskpane.html" in xml
    assert "other-host.example" not in xml


def test_manifest_uses_request_host_when_configured_localhost(tmp_path: Path) -> None:
    staff = tmp_path / "staff.json"
    staff.write_text("[]", encoding="utf-8")
    settings = Settings(
        log_path=tmp_path / "assignletters.log",
        staff_path=staff,
        public_base_url="https://localhost:8000",
    )
    client = TestClient(create_app(settings))
    xml = client.get(
        "/manifest.xml",
        headers={
            "Host": "abc.ngrok-free.app",
            "X-Forwarded-Host": "abc.ngrok-free.app",
            "X-Forwarded-Proto": "https",
        },
    ).text
    assert "https://abc.ngrok-free.app/taskpane.html" in xml
    assert "<AppDomain>https://abc.ngrok-free.app</AppDomain>" in xml
    assert "localhost:8000" not in xml


def test_taskpane_assets(tmp_path: Path) -> None:
    client = make_client(tmp_path)
    html = client.get("/taskpane.html")
    assert html.status_code == 200
    assert "office.js" in html.text
    js = client.get("/taskpane.js")
    assert js.status_code == 200
    assert "getAllInternetHeadersAsync" in js.text
    icon = client.get("/icons/icon-64.png")
    assert icon.status_code == 200
    assert icon.content[:8] == b"\x89PNG\r\n\x1a\n"
