from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from skills import publish_review_site


def test_build_share_site_copies_html_json_and_manifest(tmp_path) -> None:
    output_dir = tmp_path / "output"
    site_dir = tmp_path / "site"
    output_dir.mkdir()

    source_html = output_dir / "april_review_interactive_clean.html"
    source_json = output_dir / "april_review_interactive_clean.json"
    plotly_asset = output_dir / "plotly-2.35.2.min.js"

    source_html.write_text("<html><body>latest review</body></html>", encoding="utf-8")
    source_json.write_text('{"gmv": 123}', encoding="utf-8")
    plotly_asset.write_text("window.Plotly={};", encoding="utf-8")

    result = publish_review_site.build_share_site(
        source_html=source_html,
        source_json=source_json,
        site_dir=site_dir,
        site_url="https://example.com/review/",
    )

    assert result["index_html"].read_text(encoding="utf-8") == "<html><body>latest review</body></html>"
    assert (site_dir / "latest.json").read_text(encoding="utf-8") == '{"gmv": 123}'
    assert (site_dir / "plotly-2.35.2.min.js").read_text(encoding="utf-8") == "window.Plotly={};"
    assert (site_dir / ".nojekyll").exists()

    manifest = json.loads((site_dir / "share_manifest.json").read_text(encoding="utf-8"))
    assert manifest["site_url"] == "https://example.com/review/"
    assert manifest["source_html"] == str(source_html.resolve())


def test_resolve_source_html_prefers_clean_output(tmp_path, monkeypatch) -> None:
    clean_html = tmp_path / "april_review_interactive_clean.html"
    normal_html = tmp_path / "april_review_interactive.html"
    clean_html.write_text("clean", encoding="utf-8")
    normal_html.write_text("normal", encoding="utf-8")

    monkeypatch.setattr(
        publish_review_site,
        "DEFAULT_SOURCE_HTML_CANDIDATES",
        [clean_html, normal_html],
    )

    assert publish_review_site.resolve_source_html() == clean_html
