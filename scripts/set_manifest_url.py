#!/usr/bin/env python3
"""Rewrite addin/manifest.xml URLs for an HTTPS sideload host."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend.main import inject_manifest_urls  # noqa: E402

TEMPLATE = ROOT / "addin" / "manifest.template.xml"
OUTPUT = ROOT / "addin" / "manifest.xml"


def main() -> None:
    parser = argparse.ArgumentParser(description="Point the Outlook manifest at an HTTPS base URL.")
    parser.add_argument("base_url", help="HTTPS origin, e.g. https://abc.ngrok-free.app")
    args = parser.parse_args()
    base = args.base_url.strip().rstrip("/")
    if not base.lower().startswith("https://"):
        raise SystemExit("Base URL must start with https://")
    xml = inject_manifest_urls(TEMPLATE.read_text(encoding="utf-8"), base)
    OUTPUT.write_text(xml, encoding="utf-8")
    print(f"Wrote {OUTPUT} with base {base}")


if __name__ == "__main__":
    main()
