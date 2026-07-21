from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from skills.build_april_review import (
    normalize_product_name,
    prepare_outer_rows,
    prepare_shop_rows,
    product_match_key,
)


def test_normalize_product_name_keeps_shop_display_name() -> None:
    assert normalize_product_name("2件套-视黄醇精华+面霜") == "2件套-视黄醇精华+面霜"
    assert normalize_product_name("3件套-VC精华+视黄醇精华+视黄醇面霜") == "3件套-VC精华+视黄醇精华+视黄醇面霜"


def test_product_match_key_groups_aliases_without_merging_three_piece_bundle() -> None:
    duo_key = product_match_key("2件套-视黄醇精华+面霜")
    assert duo_key == product_match_key("2件套-精华+面霜")
    assert duo_key == product_match_key("PDRN面霜+视黄醇精华")
    assert duo_key != product_match_key("3件套-VC精华+视黄醇精华+视黄醇面霜")
    assert product_match_key("水光气垫-链子版") == product_match_key("水光气垫")
    assert product_match_key("2件套-美白精华+面霜") == product_match_key("2件套-精华+美白面霜")
    assert product_match_key("combo:水油喷雾+素颜霜") == product_match_key("2件套-素颜霜+水油喷雾")


def test_prepare_shop_rows_preserves_distinct_shop_product_names() -> None:
    shop = pd.DataFrame(
        [
            {
                "日期": "2026-04-01",
                "产品ID": 1,
                "品类": "套组",
                "站内产品命名": "2件套-精华+面霜",
                "站外产品命名": "combo:视黄醇面霜+视黄醇精华",
                "分组情况": "护肤",
                "店铺": "旗舰店",
                "GMV-人民币": 100.0,
                "链接访客": 10,
                "订单数": 1,
            },
            {
                "日期": "2026-04-01",
                "产品ID": 2,
                "品类": "套组",
                "站内产品命名": "2件套-视黄醇精华+面霜",
                "站外产品命名": "PDRN面霜+视黄醇精华",
                "分组情况": "护肤",
                "店铺": "旗舰店",
                "GMV-人民币": 200.0,
                "链接访客": 20,
                "订单数": 2,
            },
        ]
    )
    period = {
        "prior_start": pd.Timestamp("2026-03-01"),
        "current_end": pd.Timestamp("2026-04-30"),
    }

    rows, _ = prepare_shop_rows(shop, period)
    products = sorted({row["product"] for row in rows})

    assert products == ["2件套-精华+面霜", "2件套-视黄醇精华+面霜"]


def test_prepare_outer_rows_uses_outer_name_when_inner_name_is_other() -> None:
    outer = pd.DataFrame(
        [
            {
                "Date_start": "2026-04-01",
                "Spend": 10,
                "Impressions": 100,
                "Link_Click": 5,
                "Purchase Value": 20,
                "Conversions": 1,
                "产品": "foo",
                "站外产品命名": "水光气垫-链子版",
                "站内名称": "其他",
                "店铺区分": "foo",
                "品类": "气垫",
                "站内店铺命名": "旗舰店",
                "素材类型": "foo",
                "广告类型": "foo",
                "人群": "foo",
                "分组情况": "底妆",
                "模块": "foo",
                "汇率": 1,
                "板块": "foo",
            }
        ]
    )
    period = {
        "prior_start": pd.Timestamp("2026-03-01"),
        "current_end": pd.Timestamp("2026-04-30"),
    }
    resolution = {
        ("旗舰店", product_match_key("水光气垫")): {
            "product": "水光气垫",
            "group": "底妆",
            "category": "气垫",
        }
    }

    rows, _ = prepare_outer_rows(outer, period, resolution)

    assert rows[0]["product"] == "水光气垫"
