from __future__ import annotations

from typing import Any
from urllib.parse import urlparse

import httpx

DEFAULT_SAVE_PATH = "/api/save"
DEFAULT_TIMEOUT_SECONDS = 15.0


class PushError(RuntimeError):
    """Remote save API could not be called or returned an error."""


def assignment_push_url(api_url: str) -> str | None:
    """Resolve ASSIGNLETTERS_API_URL to the POST destination.

    A full URL with a path is used as-is. An origin only (no path) receives
    POST /api/save. Empty values mean the assignment is logged locally only.
    """
    raw = (api_url or "").strip()
    if not raw:
        return None
    if "://" not in raw:
        raw = "https://" + raw.lstrip("/")
    parsed = urlparse(raw)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise PushError(f"ASSIGNLETTERS_API_URL is not a valid http(s) URL: {api_url!r}")
    origin = f"{parsed.scheme}://{parsed.netloc}"
    path = parsed.path or ""
    if path in ("", "/"):
        return origin + DEFAULT_SAVE_PATH
    url = origin + path
    if parsed.query:
        url += "?" + parsed.query
    return url


def push_saved_assignment(
    api_url: str,
    payload: dict[str, Any],
    *,
    timeout: float = DEFAULT_TIMEOUT_SECONDS,
    http_client: httpx.Client | None = None,
) -> dict[str, Any]:
    url = assignment_push_url(api_url)
    if not url:
        return {"pushed": False, "url": None}
    owns_client = http_client is None
    client = http_client or httpx.Client(timeout=timeout, follow_redirects=True)
    try:
        try:
            response = client.post(url, json=payload)
        except httpx.HTTPError as exc:
            raise PushError(f"Could not reach save API at {url}: {exc}") from exc
    finally:
        if owns_client:
            client.close()
    if response.status_code >= 400:
        detail = response.text[:500] if response.text else response.reason_phrase
        raise PushError(f"Save API {url} returned HTTP {response.status_code}: {detail}")
    remote: Any = None
    try:
        remote = response.json()
    except ValueError:
        remote = None
    if isinstance(remote, dict) and remote.get("ok") is False:
        raise PushError(f"Save API {url} returned ok=false")
    return {"pushed": True, "url": url, "statusCode": response.status_code}
