from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SITE_DIR = ROOT / "site"
DEFAULT_DASHBOARD_DIR_CANDIDATES = [
    Path.home() / "Desktop" / "KOL素材分析",
    Path.home() / "OneDrive" / "Desktop" / "KOL素材分析",
]
DEFAULT_SOURCE_HTML_NAME = "SKT-合创内容数据看板_自动版.html"
DEFAULT_GENERATOR_NAME = "generate_hechuang_dashboard_from_gsheet.py"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh and publish the SKT creator dashboard to a shareable static site.")
    parser.add_argument(
        "--dashboard-dir",
        default="",
        help="Optional dashboard workspace directory. Defaults to the first existing Desktop/KOL素材分析 path.",
    )
    parser.add_argument(
        "--source-html",
        default="",
        help="Optional source HTML path. Defaults to <dashboard-dir>/SKT-合创内容数据看板_自动版.html.",
    )
    parser.add_argument(
        "--generator",
        default="",
        help="Optional generator path. Defaults to <dashboard-dir>/generate_hechuang_dashboard_from_gsheet.py.",
    )
    parser.add_argument(
        "--site-dir",
        default=str(DEFAULT_SITE_DIR),
        help="Static site output directory.",
    )
    parser.add_argument(
        "--site-url",
        default="",
        help="Optional fixed URL recorded in the manifest.",
    )
    parser.add_argument(
        "--skip-refresh",
        action="store_true",
        help="Skip running the generator and publish the current HTML directly.",
    )
    return parser.parse_args()


def resolve_dashboard_dir(override: str = "") -> Path:
    if override:
        path = Path(override)
        if not path.exists():
            raise FileNotFoundError(f"Dashboard directory not found: {path}")
        return path
    for candidate in DEFAULT_DASHBOARD_DIR_CANDIDATES:
        if candidate.exists():
            return candidate
    searched = ", ".join(str(path) for path in DEFAULT_DASHBOARD_DIR_CANDIDATES)
    raise FileNotFoundError(f"Dashboard directory not found. Searched: {searched}")


def resolve_source_html(dashboard_dir: Path, source_html: str = "") -> Path:
    if source_html:
        path = Path(source_html)
        if not path.exists():
            raise FileNotFoundError(f"Source HTML not found: {path}")
        return path
    path = dashboard_dir / DEFAULT_SOURCE_HTML_NAME
    if not path.exists():
        raise FileNotFoundError(f"Source HTML not found: {path}")
    return path


def resolve_generator(dashboard_dir: Path, generator: str = "") -> Path:
    if generator:
        path = Path(generator)
        if not path.exists():
            raise FileNotFoundError(f"Generator not found: {path}")
        return path
    path = dashboard_dir / DEFAULT_GENERATOR_NAME
    if not path.exists():
        raise FileNotFoundError(f"Generator not found: {path}")
    return path


def refresh_dashboard(generator_path: Path) -> None:
    subprocess.run(
        [sys.executable, str(generator_path)],
        cwd=str(generator_path.parent),
        check=True,
    )


def build_share_site(
    dashboard_dir: str | Path = "",
    source_html: str | Path = "",
    generator: str | Path = "",
    site_dir: str | Path = DEFAULT_SITE_DIR,
    site_url: str = "",
    skip_refresh: bool = False,
) -> dict[str, Any]:
    resolved_dashboard_dir = resolve_dashboard_dir(str(dashboard_dir) if dashboard_dir else "")
    resolved_generator = resolve_generator(resolved_dashboard_dir, str(generator) if generator else "")
    if not skip_refresh:
        refresh_dashboard(resolved_generator)

    resolved_html = resolve_source_html(resolved_dashboard_dir, str(source_html) if source_html else "")
    resolved_site_dir = Path(site_dir)
    resolved_site_dir.mkdir(parents=True, exist_ok=True)

    index_html = resolved_site_dir / "index.html"
    shutil.copyfile(resolved_html, index_html)
    (resolved_site_dir / ".nojekyll").write_text("", encoding="ascii")

    manifest = {
        "published_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "dashboard_dir": str(resolved_dashboard_dir.resolve()),
        "source_html": str(resolved_html.resolve()),
        "generator": str(resolved_generator.resolve()),
        "site_url": site_url,
        "site_entry": "index.html",
        "auto_refresh_on_publish": not skip_refresh,
    }
    manifest_path = resolved_site_dir / "share_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "site_dir": resolved_site_dir,
        "index_html": index_html,
        "manifest_path": manifest_path,
        "source_html": resolved_html,
        "generator": resolved_generator,
        "skip_refresh": skip_refresh,
    }


def main() -> None:
    args = parse_args()
    result = build_share_site(
        dashboard_dir=args.dashboard_dir,
        source_html=args.source_html,
        generator=args.generator,
        site_dir=args.site_dir,
        site_url=args.site_url,
        skip_refresh=args.skip_refresh,
    )
    print(f"Share site updated: {result['index_html']}")
    print(f"Manifest updated: {result['manifest_path']}")
    if result["skip_refresh"]:
        print("Dashboard refresh skipped. Published the current local HTML.")
    else:
        print("Dashboard data refreshed from Google Sheet before publishing.")


if __name__ == "__main__":
    main()
