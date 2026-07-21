from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from skills import build_april_review_clean as review_clean


def test_build_competitor_analysis_from_rank_csvs(tmp_path) -> None:
    main_workbook = tmp_path / "main.xlsx"
    pd.DataFrame(
        [
            {"Ranking": 1, "Product Name": "Skintific PDRN Cream", "Product Link": "a", "Shop Name": "Skintific Official", "Is it My Shop": 1},
            {"Ranking": 2, "Product Name": "BrandX Bright Serum", "Product Link": "b", "Shop Name": "BrandX Official", "Is it My Shop": 0},
        ]
    ).to_csv(tmp_path / "competitor_sales_rank.csv", index=False)
    pd.DataFrame(
        [
            {"Ranking": 1, "Product Name": "Skintific PDRN Cream", "Product Link": "a", "Shop Name": "Skintific Official", "Is it My Shop": 1},
            {"Ranking": 3, "Product Name": "BrandX Bright Serum", "Product Link": "b", "Shop Name": "BrandX Official", "Is it My Shop": 0},
        ]
    ).to_csv(tmp_path / "competitor_units_rank.csv", index=False)
    pd.DataFrame(
        [
            {"Ranking": 2, "Product Name": "Skintific PDRN Cream", "Product Link": "a", "Shop Name": "Skintific Official", "Is it My Shop": 1},
            {"Ranking": 1, "Product Name": "BrandX Bright Serum", "Product Link": "b", "Shop Name": "BrandX Official", "Is it My Shop": 0},
        ]
    ).to_csv(tmp_path / "competitor_views_rank.csv", index=False)

    config = {
        "local_cache_workbook": str(main_workbook),
        "competitor_sources": {
            "enabled": True,
            "search_dirs": [str(tmp_path)],
            "filename_keywords": ["competitor", "rank"],
            "preferred_sheet_names": {
                "sales": ["Sales"],
                "units": ["Units"],
                "views": ["Views"],
                "definitions": ["Definitions"],
            },
        },
    }

    result = review_clean.build_competitor_analysis(config=config, workbook_override=str(main_workbook))

    assert result["detected"] is True
    assert "competitor-product-analysis" in result["html"]
    assert "$skintific-competitor-product-analysis" in result["html"]
    assert "Skintific PDRN Cream" in result["html"]


def test_load_source_config_keeps_nested_competitor_sources(tmp_path) -> None:
    config_path = tmp_path / "data_sources.yaml"
    config_path.write_text(
        """default_spreadsheet_url: "x"
default_spreadsheet_key: "y"
local_cache_workbook: "output/sheet_check.xlsx"

worksheets:
  outer_ads: "A"
  shop_gmv: "B"

competitor_sources:
  enabled: true
  search_dirs:
    - "output"
  filename_keywords:
    - "competitor"
""",
        encoding="utf-8",
    )

    config = review_clean.load_source_config(config_path)

    assert isinstance(config["competitor_sources"], dict)
    assert config["competitor_sources"]["enabled"] is True
    assert config["competitor_sources"]["search_dirs"] == ["output"]


def test_render_html_appends_competitor_section() -> None:
    html = review_clean.render_html(
        payload={
            "meta": {"title": "Test Review"},
            "competitor_analysis": {
                "detected": True,
                "html": '<section id="competitor-product-analysis">Competitor Block</section>',
            },
        }
    )

    assert "competitor-product-analysis" in html
    assert "Competitor Block" in html


def test_competitor_own_brand_detection_ignores_conflicting_my_shop_flag() -> None:
    assert review_clean._competitor_is_skintific(
        product_name="Glad2Glow Brightening Lip Serum",
        shop_name="Glad2Glow Official Store",
        is_my_shop=True,
    ) is False
    assert review_clean._competitor_is_skintific(
        product_name="",
        shop_name="",
        is_my_shop=True,
    ) is True
