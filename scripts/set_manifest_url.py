#!/usr/bin/env python3
"""Write addin/manifest.xml from ASSIGNLETTERS_PUBLIC_BASE_URL or a CLI origin."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from backend.config import get_settings  # noqa: E402
from backend.manifest import generate_manifest  # noqa: E402

TEMPLATE = ROOT / "addin" / "manifest.template.xml"
OUTPUT = ROOT / "addin" / "manifest.xml"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate addin/manifest.xml from ASSIGNLETTERS_PUBLIC_BASE_URL."
    )
    parser.add_argument(
        "base_url",
        nargs="?",
        help="HTTPS origin. Defaults to ASSIGNLETTERS_PUBLIC_BASE_URL from the environment.",
    )
    args = parser.parse_args()
    base = (args.base_url or get_settings().public_base_url).strip()
    if not base.lower().startswith("https://"):
        raise SystemExit("Base URL must start with https:// (set ASSIGNLETTERS_PUBLIC_BASE_URL)")
    xml = generate_manifest(TEMPLATE.read_text(encoding="utf-8"), base)
    OUTPUT.write_text(xml, encoding="utf-8")
    print(f"Wrote {OUTPUT} with base {base}")


if __name__ == "__main__":
    main()
