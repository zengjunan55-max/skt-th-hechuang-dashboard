from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SITE_DIR = ROOT / "site"
DEFAULT_SOURCE_HTML_CANDIDATES = [
    ROOT / "output" / "april_review_interactive_clean.html",
    ROOT / "output" / "april_review_interactive.html",
]
DEFAULT_SOURCE_JSON_CANDIDATES = [
    ROOT / "output" / "april_review_interactive_clean.json",
    ROOT / "output" / "april_review_interactive.json",
]
DEFAULT_ASSET_NAMES = ["plotly-2.35.2.min.js"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare the latest review as a shareable static site.")
    parser.add_argument(
        "--source-html",
        default="",
        help="Optional source HTML path. Defaults to the latest known review output in output/.",
    )
    parser.add_argument(
        "--source-json",
        default="",
        help="Optional source JSON path. Defaults to the matching JSON file in output/ when present.",
    )
    parser.add_argument(
        "--site-dir",
        default=str(DEFAULT_SITE_DIR),
        help="Static site output directory. This directory can be deployed to GitHub Pages or any static host.",
    )
    parser.add_argument(
        "--site-url",
        default="",
        help="Optional fixed URL. Stored in the manifest for reference only.",
    )
    return parser.parse_args()


def _first_existing(paths: list[Path]) -> Path | None:
    for path in paths:
        if path.exists():
            return path
    return None


def resolve_source_html(source_html: str = "") -> Path:
    if source_html:
        path = Path(source_html)
        if not path.exists():
            raise FileNotFoundError(f"Source HTML not found: {path}")
        return path
    match = _first_existing(DEFAULT_SOURCE_HTML_CANDIDATES)
    if match is None:
        searched = ", ".join(str(path) for path in DEFAULT_SOURCE_HTML_CANDIDATES)
        raise FileNotFoundError(f"No review HTML found. Searched: {searched}")
    return match


def resolve_source_json(source_json: str = "", source_html: Path | None = None) -> Path | None:
    if source_json:
        path = Path(source_json)
        if not path.exists():
            raise FileNotFoundError(f"Source JSON not found: {path}")
        return path
    if source_html is not None:
        sibling = source_html.with_suffix(".json")
        if sibling.exists():
            return sibling
    return _first_existing(DEFAULT_SOURCE_JSON_CANDIDATES)


def _copy_if_exists(source: Path, target: Path) -> Path | None:
    if not source.exists():
        return None
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, target)
    return target


def build_share_site(
    source_html: str | Path = "",
    source_json: str | Path = "",
    site_dir: str | Path = DEFAULT_SITE_DIR,
    site_url: str = "",
) -> dict[str, Any]:
    resolved_html = resolve_source_html(str(source_html) if source_html else "")
    resolved_json = resolve_source_json(str(source_json) if source_json else "", resolved_html)
    resolved_site_dir = Path(site_dir)
    resolved_site_dir.mkdir(parents=True, exist_ok=True)

    index_html = resolved_site_dir / "index.html"
    shutil.copyfile(resolved_html, index_html)

    copied_assets: list[str] = []
    for asset_name in DEFAULT_ASSET_NAMES:
        asset_source = resolved_html.parent / asset_name
        asset_target = resolved_site_dir / asset_name
        copied = _copy_if_exists(asset_source, asset_target)
        if copied is not None:
            copied_assets.append(copied.name)

    copied_json = None
    if resolved_json is not None:
        copied_json = resolved_site_dir / "latest.json"
        shutil.copyfile(resolved_json, copied_json)

    (resolved_site_dir / ".nojekyll").write_text("", encoding="ascii")

    manifest = {
        "published_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "source_html": str(resolved_html.resolve()),
        "source_json": str(resolved_json.resolve()) if resolved_json is not None else "",
        "site_url": site_url,
        "site_entry": "index.html",
        "copied_assets": copied_assets,
    }
    manifest_path = resolved_site_dir / "share_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "site_dir": resolved_site_dir,
        "index_html": index_html,
        "latest_json": copied_json,
        "manifest_path": manifest_path,
        "source_html": resolved_html,
        "source_json": resolved_json,
        "copied_assets": copied_assets,
    }


def main() -> None:
    args = parse_args()
    result = build_share_site(
        source_html=args.source_html,
        source_json=args.source_json,
        site_dir=args.site_dir,
        site_url=args.site_url,
    )
    print(f"Share site updated: {result['index_html']}")
    if result["latest_json"] is not None:
        print(f"Payload copied: {result['latest_json']}")
    print(f"Manifest updated: {result['manifest_path']}")


if __name__ == "__main__":
    main()
