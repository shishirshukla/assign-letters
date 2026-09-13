from __future__ import annotations

from urllib.parse import urlparse

DEFAULT_SAVE_PATH = "/api/save"


class PushUrlError(ValueError):
    """ASSIGNLETTERS_API_URL is missing or not a valid http(s) URL."""


def assignment_push_url(api_url: str) -> str | None:
    """Resolve ASSIGNLETTERS_API_URL for the task pane.

    A full URL with a path is used as-is. An origin only (no path) receives
    POST /api/save. Empty values mean no remote save URL is configured.
    """
    raw = (api_url or "").strip()
    if not raw:
        return None
    if "://" not in raw:
        raw = "https://" + raw.lstrip("/")
    parsed = urlparse(raw)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise PushUrlError(f"ASSIGNLETTERS_API_URL is not a valid http(s) URL: {api_url!r}")
    origin = f"{parsed.scheme}://{parsed.netloc}"
    path = parsed.path or ""
    if path in ("", "/"):
        return origin + DEFAULT_SAVE_PATH
    url = origin + path
    if parsed.query:
        url += "?" + parsed.query
    return url
