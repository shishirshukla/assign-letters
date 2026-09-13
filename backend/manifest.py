from __future__ import annotations

from urllib.parse import urlparse
from xml.etree import ElementTree as ET

from fastapi import Request

PLACEHOLDER_ORIGIN = "https://localhost:8000"


class ManifestUrlError(ValueError):
    """Raised when a manifest URL cannot be parsed as an absolute http(s) URL."""


def normalize_public_origin(base_url: str) -> str:
    """Return scheme://host[:port] with no path or trailing slash.

    Outlook on the web builds `new URL(...)` from AppDomain and DefaultValue
    attributes. A hostname without a scheme (for example `localhost`) throws
    "Failed to construct URL" when sideloading.
    """
    raw = (base_url or "").strip()
    if not raw:
        raise ManifestUrlError("Public base URL is empty")
    if "://" not in raw:
        raw = "https://" + raw.lstrip("/")
    parsed = urlparse(raw)
    if parsed.scheme not in ("http", "https"):
        raise ManifestUrlError(f"Manifest URLs must use http or https, got {parsed.scheme!r}")
    if not parsed.hostname:
        raise ManifestUrlError(f"Could not parse a host from {base_url!r}")
    origin = f"{parsed.scheme}://{parsed.hostname}"
    if parsed.port:
        origin = f"{origin}:{parsed.port}"
    return origin


def is_loopback_origin(origin: str) -> bool:
    try:
        host = (urlparse(origin).hostname or "").lower()
    except ValueError:
        return False
    return host in {"localhost", "127.0.0.1", "::1"}


def origin_from_request(request: Request) -> str:
    proto = (request.headers.get("x-forwarded-proto") or request.url.scheme).split(",")[0].strip()
    host = (
        request.headers.get("x-forwarded-host")
        or request.headers.get("host")
        or request.url.netloc
    ).split(",")[0].strip()
    return normalize_public_origin(f"{proto}://{host}")


def resolve_manifest_origin(configured: str, request: Request | None = None) -> str:
    origin = normalize_public_origin(configured)
    if request is None or not is_loopback_origin(origin):
        return origin
    incoming = origin_from_request(request)
    if not is_loopback_origin(incoming):
        # Sideload file downloaded from the public host (ngrok, Railway, …)
        # even when .env still points at localhost.
        return incoming
    return origin


def inject_manifest_urls(template: str, base_url: str) -> str:
    origin = normalize_public_origin(base_url)
    xml = template.replace(PLACEHOLDER_ORIGIN, origin)
    xml = xml.replace("http://localhost:8000", origin)
    validate_manifest_urls(xml)
    return xml


def validate_manifest_urls(xml: str) -> None:
    root = ET.fromstring(xml)
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "AppDomain":
            text = (element.text or "").strip()
            _require_absolute_url(text, context="AppDomain")
            parsed = urlparse(text)
            if parsed.path not in ("", "/"):
                raise ManifestUrlError(
                    f"AppDomain must be an origin with no path, got {text!r}"
                )
        default = element.attrib.get("DefaultValue")
        if default and tag in {"IconUrl", "HighResolutionIconUrl", "SupportUrl", "SourceLocation", "Image", "Url"}:
            _require_absolute_url(default, context=tag)


def _require_absolute_url(value: str, *, context: str) -> None:
    if not value:
        raise ManifestUrlError(f"{context} is empty")
    parsed = urlparse(value)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise ManifestUrlError(
            f"{context} is not an absolute URL ({value!r}). "
            "Outlook reports this as 'Error in reading the manifest, Failed to construct URL'."
        )
