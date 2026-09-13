from __future__ import annotations

import json
import logging
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.middleware.base import BaseHTTPMiddleware

from backend.config import ROOT, Settings, get_settings
from backend.logging_setup import configure_logging
from backend.manifest import (
    ManifestUrlError,
    generate_manifest,
    resolve_manifest_origin,
)

logger = logging.getLogger("assignletters")

ADDIN_DIR = ROOT / "addin"
MANIFEST_TEMPLATE = ADDIN_DIR / "manifest.template.xml"


class SaveAssignmentRequest(BaseModel):
    staffName: str = Field(min_length=1)
    department: str = Field(min_length=1)
    deadlineDate: str = Field(min_length=1)
    internetHeaderName: str | None = None
    internetHeaderValue: str = Field(min_length=1)
    userEmail: str | None = None
    subject: str | None = None
    itemId: str | None = None


class RequestLogMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        request_id = request.headers.get("x-request-id") or uuid.uuid4().hex[:12]
        request.state.request_id = request_id
        start = time.perf_counter()
        try:
            response = await call_next(request)
        except Exception:
            duration_ms = int((time.perf_counter() - start) * 1000)
            logger.exception(
                "unhandled_error method=%s path=%s duration_ms=%s",
                request.method,
                request.url.path,
                duration_ms,
                extra={"request_id": request_id},
            )
            raise
        duration_ms = int((time.perf_counter() - start) * 1000)
        logger.info(
            "http_request method=%s path=%s status=%s duration_ms=%s origin=%s",
            request.method,
            request.url.path,
            response.status_code,
            duration_ms,
            request.headers.get("origin", "-"),
            extra={"request_id": request_id},
        )
        response.headers["X-Request-Id"] = request_id
        return response


def load_staff(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise HTTPException(status_code=500, detail=f"Staff file not found: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail=f"Invalid staff JSON: {exc}") from exc
    if isinstance(data, dict) and "staff" in data:
        data = data["staff"]
    if not isinstance(data, list):
        raise HTTPException(status_code=500, detail="Staff JSON must be a list of objects")
    return data


def filter_staff(rows: list[dict[str, Any]], email: str | None) -> list[dict[str, Any]]:
    if not email:
        return rows
    needle = email.strip().lower()
    matched = []
    for row in rows:
        department = str(row.get("Department") or "")
        if needle == department.strip().lower() or needle in department.lower():
            matched.append(row)
    return matched


def create_app(settings: Settings | None = None) -> FastAPI:
    injected = settings

    def load_settings() -> Settings:
        return injected if injected is not None else get_settings()

    settings = load_settings()
    log_path = configure_logging(settings.log_path)
    logger.info(
        "startup header=%s log_path=%s staff_path=%s public_base_url=%s",
        settings.header_name,
        log_path,
        settings.staff_path,
        settings.public_base_url,
        extra={"request_id": "startup"},
    )

    app = FastAPI(title="AssignLetters", version="1.0.0")
    app.state.settings = settings
    app.add_middleware(RequestLogMiddleware)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origins,
        allow_credentials=False,
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=["*"],
        expose_headers=["X-Request-Id"],
    )

    @app.get("/health")
    def health() -> dict[str, Any]:
        return {
            "ok": True,
            "service": "assignletters",
            "headerName": settings.header_name,
            "publicBaseUrl": settings.public_base_url,
        }

    @app.get("/api/config")
    def api_config() -> dict[str, Any]:
        return {
            "headerName": settings.header_name,
            "missingHeaderMessage": settings.missing_header_message,
            # Empty means the task pane should call the same origin it was loaded from.
            "apiUrl": settings.api_url.rstrip("/") if settings.api_url else "",
            "publicBaseUrl": settings.public_base_url,
        }

    @app.get("/api/staff")
    def api_staff(email: str | None = Query(default=None)) -> dict[str, Any]:
        rows = load_staff(settings.staff_path)
        filtered = filter_staff(rows, email)
        logger.info(
            "staff_list email=%s total=%s matched=%s",
            email or "-",
            len(rows),
            len(filtered),
            extra={"request_id": "-"},
        )
        return {"staff": filtered, "count": len(filtered), "filteredBy": email}

    @app.post("/api/save")
    def api_save(payload: SaveAssignmentRequest, request: Request) -> dict[str, Any]:
        request_id = getattr(request.state, "request_id", "-")
        record = {
            "savedAt": datetime.now(timezone.utc).isoformat(),
            "requestId": request_id,
            **payload.model_dump(),
        }
        logger.info(
            "assignment_saved staffName=%s department=%s deadlineDate=%s "
            "headerName=%s headerValue=%s userEmail=%s subject=%s",
            payload.staffName,
            payload.department,
            payload.deadlineDate,
            payload.internetHeaderName or settings.header_name,
            payload.internetHeaderValue,
            payload.userEmail or "-",
            payload.subject or "-",
            extra={"request_id": request_id},
        )
        return {"ok": True, "status": "Success", "assignment": record}

    @app.get("/api/logs")
    def api_logs(
        format: str = Query(default="text"),
        tail: int | None = Query(default=None, ge=1),
    ):
        path: Path = settings.log_path
        if not path.exists():
            body = ""
        else:
            body = path.read_text(encoding="utf-8", errors="replace")
        lines = body.splitlines()
        if tail:
            lines = lines[-tail:]
            body = "\n".join(lines) + ("\n" if lines else "")
        logger.info(
            "logs_viewed format=%s lines=%s",
            format,
            len(lines),
            extra={"request_id": "-"},
        )
        if format == "json":
            return {
                "path": str(path),
                "count": len(lines),
                "logs": list(reversed(lines)),
            }
        return PlainTextResponse(body or "(log file is empty)\n")

    @app.get("/logs", response_class=HTMLResponse)
    def logs_page():
        path: Path = settings.log_path
        text = path.read_text(encoding="utf-8", errors="replace") if path.exists() else "(no logs yet)\n"
        escaped = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>AssignLetters logs</title>
  <style>
    body {{ font-family: Segoe UI, system-ui, sans-serif; margin: 24px; color: #201f1e; }}
    pre {{ background: #f3f2f1; padding: 16px; overflow: auto; font-size: 12px; }}
    a {{ color: #0f6cbd; }}
  </style>
</head>
<body>
  <h1>AssignLetters logs</h1>
  <p><a href="/api/logs">Raw text</a> · <a href="/api/logs?format=json">JSON</a> · <a href="/">Home</a></p>
  <pre>{escaped}</pre>
</body>
</html>"""

    @app.get("/manifest.xml")
    def manifest(request: Request):
        if not MANIFEST_TEMPLATE.exists():
            raise HTTPException(status_code=500, detail="Manifest template missing")
        try:
            current = load_settings()
            origin = resolve_manifest_origin(current.public_base_url, request)
            xml = generate_manifest(
                MANIFEST_TEMPLATE.read_text(encoding="utf-8"),
                origin,
            )
        except ManifestUrlError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        ET.fromstring(xml)
        return Response(content=xml, media_type="text/xml")

    @app.get("/api/deploy-info")
    def deploy_info() -> dict[str, Any]:
        return {
            "configuredBaseUrl": settings.public_base_url,
            "apiUrl": settings.resolved_api_url,
            "headerName": settings.header_name,
            "logPath": str(settings.log_path),
            "staffPath": str(settings.staff_path),
        }

    if ADDIN_DIR.exists():
        app.mount("/icons", StaticFiles(directory=ADDIN_DIR / "icons"), name="icons")

        @app.get("/")
        def home():
            return FileResponse(ADDIN_DIR / "index.html")

        @app.get("/taskpane.html")
        def taskpane():
            return FileResponse(ADDIN_DIR / "taskpane.html")

        @app.get("/taskpane.js")
        def taskpane_js():
            return FileResponse(ADDIN_DIR / "taskpane.js", media_type="text/javascript")

        @app.get("/taskpane.css")
        def taskpane_css():
            return FileResponse(ADDIN_DIR / "taskpane.css", media_type="text/css")

        @app.get("/commands.html")
        def commands():
            return FileResponse(ADDIN_DIR / "commands.html")

        @app.get("/commands.js")
        def commands_js():
            return FileResponse(ADDIN_DIR / "commands.js", media_type="text/javascript")

    return app


app = create_app()
