from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.request import urlopen

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE_CONFIG_PATH = ROOT / "config" / "data_sources.yaml"
DEFAULT_OUTPUT_HTML = ROOT / "output" / "april_review_interactive.html"
DEFAULT_OUTPUT_JSON = ROOT / "output" / "april_review_interactive.json"
PLOTLY_JS_NAME = "plotly-2.35.2.min.js"
PLOTLY_JS_URL = f"https://cdn.plot.ly/{PLOTLY_JS_NAME}"
DEFAULT_SOURCE_CONFIG = {
    "default_spreadsheet_url": "https://docs.google.com/spreadsheets/d/1JRCc3jxkpJAVg3X82NdktGwrXKeNSQo2BNW_Thv08As/edit?gid=111961780#gid=111961780",
    "default_spreadsheet_key": "1JRCc3jxkpJAVg3X82NdktGwrXKeNSQo2BNW_Thv08As",
    "local_cache_workbook": "output/sheet_check.xlsx",
    "worksheets": {
        "outer_ads": "站外广告",
        "shop_gmv": "店铺GMV",
        "inner_ads": "站内广告",
        "tt_sales": "TTSP销量数据源",
        "naming": "命名管理",
        "april_temp": "GPT_4月分析临时",
    },
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build interactive April Meta x Shopee review.")
    parser.add_argument(
        "--source-config",
        default=str(DEFAULT_SOURCE_CONFIG_PATH),
        help="YAML config with the fixed Google Sheet source and worksheet names.",
    )
    parser.add_argument(
        "--workbook",
        default="",
        help="Optional local workbook override. Defaults to the cache path from source-config.",
    )
    parser.add_argument(
        "--output-html",
        default=str(DEFAULT_OUTPUT_HTML),
        help="Output HTML path.",
    )
    parser.add_argument(
        "--output-json",
        default=str(DEFAULT_OUTPUT_JSON),
        help="Optional JSON payload path for debugging.",
    )
    return parser.parse_args()


def load_source_config(path: str | Path) -> dict[str, Any]:
    config = json.loads(json.dumps(DEFAULT_SOURCE_CONFIG))
    config_path = Path(path)
    if not config_path.exists():
        return config

    lines = config_path.read_text(encoding="utf-8").splitlines()
    worksheet_mode = False
    for raw_line in lines:
        line = raw_line.rstrip()
        if not line or line.lstrip().startswith("#"):
            continue
        if line.startswith("worksheets:"):
            worksheet_mode = True
            continue
        if worksheet_mode and line.startswith("  "):
            key, _, value = line.strip().partition(":")
            config["worksheets"][key.strip()] = value.strip().strip('"')
            continue
        worksheet_mode = False
        key, _, value = line.partition(":")
        if not _:
            continue
        config[key.strip()] = value.strip().strip('"')
    return config


def safe_text(value: Any, fallback: str = "未归类") -> str:
    if value is None:
        return fallback
    if isinstance(value, float) and pd.isna(value):
        return fallback
    text = str(value).strip()
    return text if text else fallback


def safe_number(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def month_key(value: pd.Timestamp) -> str:
    return value.strftime("%Y-%m-01")


def normalize_group_name(value: Any) -> str:
    return {"护肤 ": "护肤", "底妆 ": "底妆"}.get(safe_text(value), safe_text(value))


def normalize_category_name(value: Any) -> str:
    return safe_text(value)


def normalize_shop_name(value: Any) -> str:
    aliases = {
        "SKT_新旗舰店_ID_01": "新旗舰店",
        "新旗舰店": "新旗舰店",
        "旗舰店": "旗舰店",
        "NO.1": "NO.1",
        "C店": "C店",
        "新店": "新旗舰店",
        "3": "C店",
        "2": "新旗舰店",
        "1": "旗舰店",
    }
    text = safe_text(value)
    return aliases.get(text, text)


def normalize_product_name(value: Any) -> str:
    return safe_text(value)


def normalize_product_match_value(value: Any) -> str:
    text = safe_text(value)
    lowered = text.casefold().replace("combo:", "").strip()
    lowered = re.sub(r"-?链子版$", "", lowered)

    if "3件套" not in lowered and "精华" in lowered and "面霜" in lowered:
        if "美白" in lowered:
            return "2件套-美白精华+面霜"
        return "2件套-精华+面霜"
    if "水油喷雾" in lowered and "素颜霜" in lowered:
        return "2件套-水油喷雾+素颜霜"
    return lowered


def product_match_key(value: Any) -> str:
    text = normalize_product_match_value(value)
    return text.casefold()


def infer_period(shop_df: pd.DataFrame) -> dict[str, pd.Timestamp]:
    dates = pd.to_datetime(shop_df["日期"], errors="coerce").dropna()
    april_dates = dates[dates.dt.month == 4]
    if april_dates.empty:
        raise ValueError("No April dates found in 店铺GMV sheet.")
    current_end = april_dates.max().normalize()
    current_start = pd.Timestamp(year=current_end.year, month=4, day=1)
    day_count = (current_end - current_start).days + 1
    prior_start = current_start - pd.DateOffset(months=1)
    prior_end = prior_start + pd.Timedelta(days=day_count - 1)
    return {
        "current_start": current_start,
        "current_end": current_end,
        "prior_start": prior_start,
        "prior_end": prior_end,
        "day_count": day_count,
    }


def build_google_sheet_export_url(config: dict[str, Any]) -> str:
    key = safe_text(config.get("default_spreadsheet_key"), fallback="")
    if not key:
        raise ValueError("Missing default_spreadsheet_key in source config.")
    return f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx"


def refresh_online_workbook_cache(config: dict[str, Any]) -> Path:
    workbook_path = ROOT / config["local_cache_workbook"]
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    export_url = build_google_sheet_export_url(config)
    with urlopen(export_url) as response:
        workbook_bytes = response.read()
    if not workbook_bytes.startswith(b"PK"):
        raise RuntimeError("Downloaded workbook is not a valid XLSX file.")
    workbook_path.write_bytes(workbook_bytes)
    return workbook_path


def load_workbook_frames(config: dict[str, Any], workbook_override: str = "") -> dict[str, pd.DataFrame]:
    workbook_path = Path(workbook_override) if workbook_override else ROOT / config["local_cache_workbook"]
    if not workbook_path.exists():
        raise FileNotFoundError(
            f"Workbook not found: {workbook_path}. "
            "Please sync the fixed Google Sheet to the local cache workbook first."
        )

    sheets = config["worksheets"]
    outer = pd.read_excel(
        workbook_path,
        sheet_name=sheets["outer_ads"],
        usecols=[
            "Date_start",
            "Spend",
            "Impressions",
            "Link_Click",
            "Purchase Value",
            "Conversions",
            "产品",
            "站外产品命名",
            "站内名称",
            "店铺区分",
            "站内店铺命名",
            "品类",
            "分组情况",
            "模块",
            "汇率",
            "板块",
            "广告类型",
            "人群",
            "素材类型",
        ],
    )
    shop = pd.read_excel(
        workbook_path,
        sheet_name=sheets["shop_gmv"],
        usecols=[
            "日期",
            "产品ID",
            "品类",
            "站内产品命名",
            "站外产品命名",
            "分组情况",
            "店铺",
            "GMV-人民币",
            "链接访客",
            "订单数",
        ],
    )
    inner = pd.read_excel(
        workbook_path,
        sheet_name=sheets["inner_ads"],
        usecols=[
            "date",
            "产品ID",
            "类目",
            "站内产品命名",
            "站外产品命名",
            "Spend-人民币",
        ],
    )
    tt = pd.read_excel(
        workbook_path,
        sheet_name=sheets["tt_sales"],
        usecols="A:H",
    )
    temp = pd.read_excel(
        workbook_path,
        sheet_name=sheets["april_temp"],
        usecols="A:U",
        nrows=45,
    )
    return {"outer": outer, "shop": shop, "inner": inner, "tt": tt, "temp": temp}


def prepare_shop_rows(shop: pd.DataFrame, period: dict[str, pd.Timestamp]) -> tuple[list[dict[str, Any]], dict[Any, dict[str, str]]]:
    df = shop.copy()
    df["date"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df[df["date"].between(period["prior_start"], period["current_end"])].copy()
    df["shop"] = df["店铺"].apply(normalize_shop_name)
    df["group"] = df["分组情况"].apply(normalize_group_name)
    df["category"] = df["品类"].apply(normalize_category_name)
    df["product"] = df["站内产品命名"].apply(normalize_product_name)
    df["gmv"] = pd.to_numeric(df["GMV-人民币"], errors="coerce").fillna(0.0)
    df["visitors"] = pd.to_numeric(df["链接访客"], errors="coerce").fillna(0.0)
    df["orders"] = pd.to_numeric(df["订单数"], errors="coerce").fillna(0.0)

    grouped = (
        df.groupby(["date", "shop", "group", "category", "product"], dropna=False)[["gmv", "visitors", "orders"]]
        .sum()
        .reset_index()
    )
    grouped["day"] = grouped["date"].dt.day

    lookup_source = (
        df.sort_values("date")
        .dropna(subset=["产品ID"])
        .drop_duplicates(subset=["产品ID"], keep="last")
        [["产品ID", "shop", "group", "category", "product"]]
    )
    product_lookup: dict[Any, dict[str, str]] = {}
    for _, row in lookup_source.iterrows():
        product_lookup[row["产品ID"]] = {
            "shop": row["shop"],
            "group": row["group"],
            "category": row["category"],
            "product": row["product"],
        }

    rows = [
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "day": int(row.day),
            "shop": row.shop,
            "group": row.group,
            "category": row.category,
            "product": row.product,
            "gmv": round(float(row.gmv), 2),
            "visitors": int(row.visitors),
            "orders": int(row.orders),
        }
        for row in grouped.itertuples(index=False)
    ]
    return rows, product_lookup


def build_shop_product_resolution(shop_rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, str]]:
    if not shop_rows:
        return {}

    df = pd.DataFrame(shop_rows).copy()
    df["match_key"] = df["product"].apply(product_match_key)
    grouped = (
        df.groupby(["shop", "match_key", "product", "group", "category"], dropna=False)[["gmv"]]
        .sum()
        .reset_index()
        .sort_values(["shop", "match_key", "gmv", "product"], ascending=[True, True, False, True])
    )

    resolution: dict[tuple[str, str], dict[str, str]] = {}
    for row in grouped.itertuples(index=False):
        key = (row.shop, row.match_key)
        if key not in resolution:
            resolution[key] = {
                "product": row.product,
                "group": row.group,
                "category": row.category,
            }
    return resolution


def prepare_inner_rows(
    inner: pd.DataFrame,
    product_lookup: dict[Any, dict[str, str]],
    period: dict[str, pd.Timestamp],
) -> list[dict[str, Any]]:
    df = inner.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["date"].between(period["prior_start"], period["current_end"])].copy()
    df["product"] = df["站内产品命名"].apply(normalize_product_name)
    df["spend"] = pd.to_numeric(df["Spend-人民币"], errors="coerce").fillna(0.0)
    df["category"] = df["类目"].apply(normalize_category_name)
    df["shop"] = "未归类"
    df["group"] = "未归类"

    for idx, row in df.iterrows():
        lookup = product_lookup.get(row["产品ID"])
        if lookup:
            df.at[idx, "shop"] = lookup["shop"]
            df.at[idx, "group"] = lookup["group"]
            if lookup["product"] != "未归类":
                df.at[idx, "product"] = lookup["product"]
            if df.at[idx, "category"] == "未归类":
                df.at[idx, "category"] = lookup["category"]

    grouped = (
        df.groupby(["date", "shop", "group", "category", "product"], dropna=False)[["spend"]]
        .sum()
        .reset_index()
    )
    grouped["day"] = grouped["date"].dt.day

    return [
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "day": int(row.day),
            "shop": row.shop,
            "group": row.group,
            "category": row.category,
            "product": row.product,
            "inner_spend": round(float(row.spend), 2),
        }
        for row in grouped.itertuples(index=False)
    ]


def prepare_outer_rows(
    outer: pd.DataFrame,
    period: dict[str, pd.Timestamp],
    shop_product_resolution: dict[tuple[str, str], dict[str, str]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    df = outer.copy()
    df["date"] = pd.to_datetime(df["Date_start"], errors="coerce")
    df = df[df["date"].between(period["prior_start"], period["current_end"])].copy()
    df["shop"] = df["站内店铺命名"].apply(normalize_shop_name)
    df["group"] = df["分组情况"].apply(normalize_group_name)
    df["category"] = df["品类"].apply(normalize_category_name)
    df["product"] = df["站内名称"].apply(normalize_product_name)
    fallback_mask = df["product"].isin(["未归类", "其他"])
    df.loc[fallback_mask, "product"] = df.loc[fallback_mask, "站外产品命名"].apply(normalize_product_name)
    df["match_key"] = df["product"].apply(product_match_key)
    for idx, row in df.iterrows():
        resolved = shop_product_resolution.get((row["shop"], row["match_key"]))
        if resolved:
            df.at[idx, "product"] = resolved["product"]
            df.at[idx, "group"] = resolved["group"]
            df.at[idx, "category"] = resolved["category"]
    df["module"] = df["模块"].apply(safe_text)
    df["board"] = df["板块"].apply(safe_text)
    df["fx_rate"] = pd.to_numeric(df["汇率"], errors="coerce").fillna(0.0)
    df.loc[df["fx_rate"] <= 0, "fx_rate"] = 1.0
    df["outer_spend"] = pd.to_numeric(df["Spend"], errors="coerce").fillna(0.0) * df["fx_rate"]
    df["impressions"] = pd.to_numeric(df["Impressions"], errors="coerce").fillna(0.0)
    df["clicks"] = pd.to_numeric(df["Link_Click"], errors="coerce").fillna(0.0)
    df["ad_gmv"] = pd.to_numeric(df["Purchase Value"], errors="coerce").fillna(0.0) * df["fx_rate"]
    df["ad_orders"] = pd.to_numeric(df["Conversions"], errors="coerce").fillna(0.0)

    product_level = (
        df.groupby(["date", "shop", "group", "category", "product"], dropna=False)[
            ["outer_spend", "impressions", "clicks", "ad_gmv", "ad_orders"]
        ]
        .sum()
        .reset_index()
    )
    product_level["day"] = product_level["date"].dt.day

    module_level = (
        df.groupby(["date", "shop", "group", "category", "product", "module", "board"], dropna=False)[
            ["outer_spend", "impressions", "clicks", "ad_gmv", "ad_orders"]
        ]
        .sum()
        .reset_index()
    )
    module_level["day"] = module_level["date"].dt.day

    product_rows = [
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "day": int(row.day),
            "shop": row.shop,
            "group": row.group,
            "category": row.category,
            "product": row.product,
            "outer_spend": round(float(row.outer_spend), 2),
            "impressions": int(row.impressions),
            "clicks": int(row.clicks),
            "ad_gmv": round(float(row.ad_gmv), 2),
            "ad_orders": int(row.ad_orders),
        }
        for row in product_level.itertuples(index=False)
    ]
    module_rows = [
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "day": int(row.day),
            "shop": row.shop,
            "group": row.group,
            "category": row.category,
            "product": row.product,
            "module": row.module,
            "board": row.board,
            "outer_spend": round(float(row.outer_spend), 2),
            "impressions": int(row.impressions),
            "clicks": int(row.clicks),
            "ad_gmv": round(float(row.ad_gmv), 2),
            "ad_orders": int(row.ad_orders),
        }
        for row in module_level.itertuples(index=False)
    ]
    return product_rows, module_rows


def prepare_tt_rows(tt: pd.DataFrame, period: dict[str, pd.Timestamp]) -> list[dict[str, Any]]:
    # load_workbook_frames 读取 usecols 后会按工作表原始列顺序返回：
    # A 时间、D 总销量、E 渠道、F 品类、G 产品名、H 分组
    # 这里继续按这个固定顺序处理，确保产品口径以源表 G 列为准
    df = tt.iloc[:, [0, 3, 4, 5, 6, 7]].copy()
    df.columns = ["time_raw", "sales_raw", "channel_raw", "category_raw", "product_raw", "group_raw"]
    df["time"] = pd.to_datetime(df["time_raw"], errors="coerce")
    df["month"] = df["time"].dt.to_period("M").dt.to_timestamp()
    month_start = period["prior_start"].replace(day=1)
    month_end = period["current_end"].replace(day=1)
    df = df[df["month"].between(month_start, month_end)].copy()
    df["group"] = df["group_raw"].apply(normalize_group_name)
    df["category"] = df["category_raw"].apply(normalize_category_name)
    df["product"] = df["product_raw"].apply(safe_text)
    df["channel"] = df["channel_raw"].apply(safe_text)
    df["sales"] = pd.to_numeric(df["sales_raw"], errors="coerce").fillna(0.0)

    grouped = (
        df.groupby(["month", "group", "category", "product", "channel"], dropna=False)[["sales"]]
        .sum()
        .reset_index()
    )
    pivot = (
        grouped.pivot_table(
            index=["month", "group", "category", "product"],
            columns="channel",
            values="sales",
            fill_value=0.0,
            aggfunc="sum",
        )
        .reset_index()
        .rename_axis(None, axis=1)
    )

    tt_column = "TT" if "TT" in pivot.columns else None
    shopee_column = "Shopee" if "Shopee" in pivot.columns else None
    rows = []
    for row in pivot.itertuples(index=False):
        rows.append(
            {
                "month": month_key(row.month),
                "group": row.group,
                "category": row.category,
                "product": row.product,
                "tt_sales": int(getattr(row, tt_column, 0) if tt_column else 0),
                "shopee_sales": int(getattr(row, shopee_column, 0) if shopee_column else 0),
            }
        )
    return rows


def extract_temp_overview(temp: pd.DataFrame, period: dict[str, pd.Timestamp]) -> list[dict[str, Any]]:
    df = temp.copy()
    first_col = df.columns[0]
    df = df.rename(
        columns={
            first_col: "date",
            "GMV_RMB": "gmv",
            "Visitors": "visitors",
            "Orders": "orders",
            "OuterSpend": "outer_spend",
            "Impressions": "impressions",
            "LPV": "lpv",
            "AdGMV": "ad_gmv",
            "AdOrders": "ad_orders",
            "InnerSpend": "inner_spend",
            "Shopee": "shopee_sales",
            "TT": "tt_sales",
        }
    )
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["date"].between(period["prior_start"], period["current_end"])].copy()
    for column in ["gmv", "visitors", "orders", "outer_spend", "impressions", "lpv", "ad_gmv", "ad_orders", "inner_spend", "shopee_sales", "tt_sales"]:
        if column not in df.columns:
            df[column] = 0.0
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    df["day"] = df["date"].dt.day
    return [
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "day": int(row.day),
            "gmv": round(float(row.gmv), 2),
            "visitors": int(row.visitors),
            "orders": int(row.orders),
            "outer_spend": round(float(row.outer_spend), 2),
            "impressions": int(row.impressions),
            "lpv": int(row.lpv),
            "ad_gmv": round(float(row.ad_gmv), 2),
            "ad_orders": int(row.ad_orders),
            "inner_spend": round(float(row.inner_spend), 2),
            "shopee_sales": int(row.shopee_sales),
            "tt_sales": int(row.tt_sales),
        }
        for row in df.itertuples(index=False)
    ]


def build_filters(rows_by_source: dict[str, list[dict[str, Any]]]) -> dict[str, list[str]]:
    fields = ("shop", "group", "category", "product")
    values: dict[str, set[str]] = {field: set() for field in fields}
    for rows in rows_by_source.values():
        for row in rows:
            for field in fields:
                value = safe_text(row.get(field))
                if value != "未归类":
                    values[field].add(value)
    return {field: sorted(items) for field, items in values.items()}


def build_payload(config: dict[str, Any], workbook_override: str = "") -> dict[str, Any]:
    refreshed_workbook = None
    if workbook_override:
        workbook_path = Path(workbook_override)
    else:
        refreshed_workbook = refresh_online_workbook_cache(config)
        workbook_path = refreshed_workbook

    frames = load_workbook_frames(config=config, workbook_override=str(workbook_path))
    period = infer_period(frames["shop"])
    shop_rows, product_lookup = prepare_shop_rows(frames["shop"], period)
    shop_product_resolution = build_shop_product_resolution(shop_rows)
    inner_rows = prepare_inner_rows(frames["inner"], product_lookup, period)
    outer_rows, outer_module_rows = prepare_outer_rows(frames["outer"], period, shop_product_resolution)
    tt_rows = prepare_tt_rows(frames["tt"], period)
    overview_rows = extract_temp_overview(frames["temp"], period)
    filters = build_filters({"shop": shop_rows, "inner": inner_rows, "outer": outer_rows, "tt": tt_rows})
    return {
        "meta": {
            "title": "Skintific Indonesia Meta x Shopee 4月复盘",
            "subtitle": "Meta站外投放 x 店铺GMV x 站内广告联动复盘",
            "current_start": period["current_start"].strftime("%Y-%m-%d"),
            "current_end": period["current_end"].strftime("%Y-%m-%d"),
            "prior_start": period["prior_start"].strftime("%Y-%m-%d"),
            "prior_end": period["prior_end"].strftime("%Y-%m-%d"),
            "day_count": period["day_count"],
            "data_refresh_mode": "online_refresh" if refreshed_workbook else "workbook_override",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_rule_note": "每次生成均强制重读在线表；TT 销量按 TTSP销量数据源 G列产品名；站外 Spend/广告GMV 先按表内汇率换算人民币；站外店铺区分取 X列。",
        },
        "filters": filters,
        "shop_daily": shop_rows,
        "inner_daily": inner_rows,
        "outer_daily": outer_rows,
        "outer_module_daily": outer_module_rows,
        "tt_monthly": tt_rows,
        "overview_daily": overview_rows,
        "focus_bundles": [
            "2件套-377精华+面霜",
            "2件套-精华+面霜",
            "PDRN+素颜霜2件套",
            "2件套-磨皮棉片+美白面膜",
        ],
    }


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>__TITLE__</title>
  <script src="__PLOTLY_SRC__"></script>
  <style>
    :root{--bg:#f5f7fb;--card:#ffffff;--line:#dbe3ef;--text:#1f2937;--muted:#64748b;--blue:#3b82f6;--blue-deep:#1d4ed8;--blue-soft:rgba(59,130,246,.10);--blue-soft-2:#f8fbff;--warn:#dc2626;--shadow:0 10px 30px rgba(15,23,42,.06)}
    *{box-sizing:border-box} body{margin:0;font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif;color:var(--text);background:var(--bg)}
    .page{width:min(1440px,calc(100vw - 24px));margin:0 auto;padding:20px 0 40px}
    .hero,.card,.table-card,.filter-panel{background:rgba(255,255,255,.92);border:1px solid var(--line);border-radius:22px;box-shadow:var(--shadow)}
    .hero{padding:28px;margin-bottom:18px;background:radial-gradient(circle at top right,rgba(59,130,246,.10),transparent 32%),linear-gradient(135deg,rgba(255,255,255,.98),rgba(248,251,255,.98))}
    .filter-panel{padding:22px 24px;margin-bottom:20px;background:#eef2f8;color:var(--text);position:static}
    .summary-grid,.chart-grid,.section-grid,.note-grid{display:grid;gap:18px}
    .summary-grid{grid-template-columns:repeat(6,minmax(0,1fr));gap:14px;margin:4px 0 18px}
    .chart-grid{grid-template-columns:repeat(2,minmax(0,1fr));align-items:start;margin-bottom:18px}
    .section-grid,.note-grid{grid-template-columns:repeat(auto-fit,minmax(320px,1fr));margin-bottom:18px}
    .card,.table-card{padding:22px}
    .hero-head,.chart-head,.table-head,.filter-actions{display:flex;gap:12px;justify-content:space-between;align-items:flex-start;flex-wrap:wrap}
    .hero-head>*{min-width:0}
    .title-wrap{flex:1 1 760px;min-width:0}
    .chart-grid>.card{display:grid;grid-template-rows:auto auto auto;gap:18px;padding:26px 24px}
    .chart-head{gap:18px;margin-bottom:0}
    .chart-head>div:first-child{flex:1 1 auto;min-width:0}
    .chart-head h2{margin:10px 0 10px;font-size:clamp(22px,2.4vw,30px);line-height:1.22;color:#183764}
    .chart-head .muted{display:block;max-width:100%;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:14px;line-height:1.4}
    .title-wrap h1{margin:0 0 12px;font-size:clamp(28px,4vw,42px);line-height:1.08}
    .title-wrap p,.muted,p,li{line-height:1.75;color:var(--muted)}
    .meta-note,.chart-tools,.filter-actions{display:flex;gap:8px;flex-wrap:wrap}
    .badge,.section-label,.pill{display:inline-flex;align-items:center;padding:6px 11px;border-radius:999px;font-size:12px;font-weight:700;letter-spacing:.05em}
    .badge,.section-label{background:var(--blue-soft);color:var(--blue-deep)} .pill{background:rgba(255,255,255,.86);border:1px solid rgba(47,107,255,.14)}
    .filter-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:18px;align-items:end}
    .filter-box label,.table-config label{display:block;margin-bottom:8px;font-size:12px;color:var(--muted)}
    .filter-box select,.filter-box button,.table-config input,button{font:inherit}
    .filter-box select,.table-config input{width:100%;height:64px;border-radius:20px;border:1px solid rgba(47,107,255,.16);padding:0 20px;background:#fff;color:var(--text);box-shadow:inset 0 1px 0 rgba(255,255,255,.9)}
    button{cursor:pointer;transition:transform .16s ease,opacity .16s ease}
    button:hover{transform:translateY(-1px)}
    .btn-primary,.mini-btn{border:none;background:linear-gradient(135deg,var(--blue),#6d95ff);color:#fff;border-radius:999px;padding:10px 16px;font-weight:700}
    .btn-secondary{background:rgba(255,255,255,.82);color:var(--blue-deep);border:1px solid rgba(47,107,255,.18);border-radius:999px;padding:10px 16px}
    .chart-tools button,
    .shop-block-head button,
    button[data-download],
    button[data-shop-download]{
      min-height:40px;
      padding:0 16px;
      border-radius:16px;
      border:1px solid rgba(46,90,167,.16);
      background:#ffffff;
      color:var(--blue-deep);
      font-size:12px;
      font-weight:700;
      line-height:1;
      box-shadow:0 6px 18px rgba(46,90,167,.06), inset 0 1px 0 rgba(255,255,255,.9);
      white-space:nowrap;
    }
    .chart-tools button:hover,
    .shop-block-head button:hover,
    button[data-download]:hover,
    button[data-shop-download]:hover{
      background:#f8fbff;
      border-color:rgba(46,90,167,.22);
    }
    .kpi-value{margin:10px 0 4px;font-size:clamp(24px,3vw,34px);font-weight:700}
    .delta{font-size:13px;font-weight:700}.up{color:#15a05d}.down{color:var(--warn)}.flat{color:var(--muted)}
    .summary-card{position:relative;padding:14px 16px 13px;min-height:118px;border-radius:20px;background:#fff;border:1px solid rgba(59,130,246,.10);box-shadow:var(--shadow)}
    .summary-card .section-label{padding:0;background:transparent;border-radius:0;font-size:12px;font-weight:600;letter-spacing:0;color:#6f84a3}
    .summary-card .kpi-value{margin:7px 0 6px;font-size:clamp(20px,2.1vw,28px);line-height:1.15;color:#16345f}
    .summary-card .delta{display:block;margin-bottom:6px;font-size:12px}
    .summary-card .muted{display:block;font-size:12px;line-height:1.45;color:#7f92af}
    #source-note{flex:0 1 420px!important;max-width:min(100%,420px)!important;min-width:0;overflow-wrap:anywhere;word-break:break-word;white-space:normal}
    #source-note a{display:block;max-width:100%;overflow-wrap:anywhere;word-break:break-all}
    .plot,canvas{display:block;width:100%;min-height:360px;border-radius:18px;background:linear-gradient(180deg,rgba(47,107,255,.04),rgba(255,255,255,.96)),#fff;border:1px solid rgba(47,107,255,.10);overflow:hidden}
    .plot .modebar{top:10px!important;right:10px!important}
    .narrative,.table-config,.table-wrap{margin-top:12px;padding:16px;border-radius:18px;background:var(--blue-soft-2);border:1px solid rgba(47,107,255,.10)}
    .chart-grid .narrative{margin-top:0;padding:20px 22px;border-radius:20px;background:#fbfcff;border:1px solid var(--line)}
    .chart-grid .narrative p{margin:0 0 16px;font-size:14px;line-height:1.9}
    .chart-grid .narrative p:last-child{margin-bottom:0}
    .chart-grid .narrative strong{display:block;margin-bottom:8px;color:#1b3b6c}
    .chart-grid .narrative ul{margin:0;padding-left:22px}
    .chart-grid .narrative li{margin-bottom:10px;font-size:14px;line-height:1.9}
    .chart-grid .narrative li:last-child{margin-bottom:0}
    .narrative ul{margin:0;padding-left:18px}
    .table-config{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px}
    .toggle-chip{display:flex;gap:8px;align-items:center;padding:10px 12px;border:1px solid rgba(47,107,255,.14);border-radius:14px;background:#fff}
    .toggle-chip input[type='checkbox']{width:auto;margin:0}
    .table-wrap{overflow:auto;padding:0;background:#fff}
    table{width:100%;min-width:1100px;border-collapse:collapse;font-size:13px}
    th,td{padding:11px 10px;border-bottom:1px solid rgba(47,107,255,.10);text-align:left;vertical-align:top;white-space:nowrap}
    th{position:sticky;top:0;background:#f1f6ff;z-index:1}
    .empty-state{padding:16px;border-radius:16px;background:rgba(47,107,255,.06);color:var(--blue-deep);border:1px dashed rgba(47,107,255,.25)}
    a{color:var(--blue);text-decoration:none}
    .filter-head{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:14px}
    .filter-title{font-size:14px;font-weight:700;color:var(--blue-deep)}
    .filter-summary-line{font-size:12px;color:var(--muted)}
    @media (max-width:1360px){.summary-grid{grid-template-columns:repeat(3,minmax(0,1fr));}}
    @media (max-width:1200px){.filter-grid{grid-template-columns:repeat(3,minmax(0,1fr));}.chart-grid{grid-template-columns:1fr;}}
    @media (max-width:900px){.page{width:min(100vw - 16px,1440px)}.hero,.card,.table-card,.filter-panel{padding:18px}.filter-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.summary-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.chart-grid>.card{padding:20px 18px}.chart-head .muted{max-width:100%}}
    @media (max-width:640px){.filter-grid,.summary-grid{grid-template-columns:1fr}}
  </style>
</head>
<body>
  <div class="page">
    <section class="hero">
      <div class="hero-head">
        <div class="title-wrap">
          <div class="badge">APRIL REVIEW</div>
          <h1 id="page-title"></h1>
          <p id="page-subtitle"></p>
          <div class="meta-note">
            <span class="pill" id="period-pill"></span>
            <span class="pill" id="compare-pill"></span>
            <span class="pill">数据源固定为项目主表</span>
          </div>
        </div>
        <div class="muted" style="display:none;max-width:380px" id="source-note"></div>
      </div>
    </section>
    <section class="filter-panel">
      <div class="filter-grid">
        <div class="filter-box"><label for="shop-filter">店铺筛选</label><select id="shop-filter" multiple></select></div>
        <div class="filter-box"><label for="group-filter">分组筛选</label><select id="group-filter" multiple></select></div>
        <div class="filter-box"><label for="category-filter">品类筛选</label><select id="category-filter" multiple></select></div>
        <div class="filter-box"><label for="product-filter">产品筛选</label><select id="product-filter" multiple></select></div>
      </div>
      <div class="filter-actions">
        <button class="btn-primary" id="reset-filter">重置全部筛选</button>
        <button class="btn-secondary" id="focus-bundle">只看重点新套组</button>
        <span class="muted" id="filter-summary" style="color:rgba(248,244,238,.82)"></span>
      </div>
    </section>
    <section class="summary-grid" id="summary-grid"></section>
    <section class="chart-grid">
      <div class="card"><div class="chart-head"><div><div class="section-label">Executive Summary</div><h2>整体趋势</h2><div class="muted">店铺 GMV 与上月同期趋势对比</div></div><div class="chart-tools"><button data-download="gmv-trend">下载图表</button></div></div><canvas id="gmv-trend" width="920" height="360"></canvas><div class="narrative" id="overall-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Spend Mix</div><h2>站内外花费与店铺结果</h2><div class="muted">组合图：外投花费、站内花费、店铺 GMV</div></div><div class="chart-tools"><button data-download="spend-mix">下载图表</button></div></div><canvas id="spend-mix" width="920" height="360"></canvas><div class="narrative" id="spend-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Group</div><h2>分组分析</h2><div class="muted">看 GMV 拉动和效率变化</div></div><div class="chart-tools"><button data-download="group-chart">下载图表</button></div></div><canvas id="group-chart" width="920" height="360"></canvas><div class="narrative" id="group-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Module</div><h2>模块分析</h2><div class="muted">模块维度使用站外广告口径</div></div><div class="chart-tools"><button data-download="module-chart">下载图表</button></div></div><canvas id="module-chart" width="920" height="360"></canvas><div class="narrative" id="module-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Category</div><h2>品类分析</h2><div class="muted">GMV、访客、CVR 与站内外花费联动</div></div><div class="chart-tools"><button data-download="category-chart">下载图表</button></div></div><canvas id="category-chart" width="920" height="360"></canvas><div class="narrative" id="category-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Product</div><h2>产品分析</h2><div class="muted">Top 产品看规模、增长弹性和效率拖累</div></div><div class="chart-tools"><button data-download="product-chart">下载图表</button></div></div><canvas id="product-chart" width="920" height="360"></canvas><div class="narrative" id="product-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">TT Gap</div><h2>TT 销量 GAP</h2><div class="muted">对比 TT 与 Shopee 销量</div></div><div class="chart-tools"><button data-download="tt-gap-chart">下载图表</button></div></div><canvas id="tt-gap-chart" width="920" height="360"></canvas><div class="narrative" id="tt-gap-narrative"></div></div>
      <div class="card"><div class="chart-head"><div><div class="section-label">Bundle Focus</div><h2>重点新套组专项</h2><div class="muted">重点看 2 件套，尤其是 377 combo 与新组合验证</div></div><div class="chart-tools"><button data-download="bundle-chart">下载图表</button></div></div><canvas id="bundle-chart" width="920" height="360"></canvas><div class="analysis-hidden-legacy"><div id="bundle-box"></div></div><div class="narrative" id="bundle-narrative"></div></div>
    </section>
    <section class="section-grid">
      <div class="card"><div class="section-label">Highlights</div><h2>亮点</h2><div class="narrative" id="highlight-box"></div></div>
      <div class="card"><div class="section-label">Weaknesses</div><h2>不足</h2><div class="narrative" id="weakness-box"></div></div>
    </section>
    <section class="table-card"><div class="table-head"><div><div class="section-label">Diagnosis Table</div><h2>品类诊断表</h2><div class="muted">支持列显示控制、表头修改和行隐藏</div></div></div><div id="category-table-controls"></div><div id="category-table"></div></section>
    <section class="table-card" style="margin-top:18px"><div class="table-head"><div><div class="section-label">Diagnosis Table</div><h2>产品诊断表</h2><div class="muted">产品列统一使用站内命名</div></div></div><div id="product-table-controls"></div><div id="product-table"></div></section>
    <section class="note-grid" style="margin-top:18px">
      <div class="card"><div class="section-label">Action List</div><h2>下一步 To Do</h2><div class="narrative" id="action-box"></div></div>
      <div class="card"><div class="section-label">Data Notes</div><h2>数据说明</h2><div class="narrative" id="data-note-box"></div></div>
    </section>
  </div>
  <script>
    const DATA = __REPORT_DATA__;
    const FILTER_FIELDS = ["shop","group","category","product"];
    const METRIC_OPTIONS = [
      {value:"gmv",label:"GMV"},
      {value:"orders",label:"订单量"},
      {value:"visitors",label:"访客"},
      {value:"cvr",label:"转化率"},
      {value:"innerSpend",label:"站内花费"},
      {value:"outerSpend",label:"站外花费"}
    ];
    const TOP_N_OPTIONS = [5,8,10,15,20,30,40,50];
    const TABLE_COLUMNS = {
      category:[["label","品类"],["gmv","GMV"],["visitors","访客"],["cvr","转化率"],["innerSpend","站内花费"],["outerSpend","站外花费"],["impressions","站外曝光量"],["gmvMom","GMV环比上月同期"],["visitorsMom","访客环比上月同期"],["cvrMom","CVR环比上月同期"],["diagnosis","诊断"],["todo","下一步"]],
      product:[["label","产品"],["gmv","GMV"],["visitors","访客"],["cvr","转化率"],["innerSpend","站内花费"],["outerSpend","站外花费"],["impressions","站外曝光量"],["gmvMom","GMV环比上月同期"],["visitorsMom","访客环比上月同期"],["cvrMom","CVR环比上月同期"],["diagnosis","诊断"],["todo","下一步"]]
    };
    const state = {filters:{shop:[],group:[],category:[],product:[]},metricFocus:"gmv",topN:8,tableState:{category:initTableState("category"),product:initTableState("product")}};

    function initTableState(kind){const labels={},visible={};TABLE_COLUMNS[kind].forEach(([k,l])=>{labels[k]=l;visible[k]=true;});return {labels,visible,hiddenRows:new Set()};}
    function safeNum(v){return Number(v||0)}
    function ratio(a,b){return b?safeNum(a)/safeNum(b):0}
    function pctChange(current,prior){return !prior ? (current ? 1 : 0) : (current-prior)/prior}
    function fmtNum(v,d=0){return safeNum(v).toLocaleString("zh-CN",{minimumFractionDigits:d,maximumFractionDigits:d})}
    function fmtMoney(v){return fmtNum(v,0)}
    function fmtPct(v){return `${(safeNum(v)*100).toFixed(1)}%`}
    function fmtDelta(v){const n=safeNum(v);return `${n>0?"+":""}${(n*100).toFixed(1)}%`}
    function deltaClass(v){if(v>.005)return"up";if(v<-.005)return"down";return"flat"}
    function narrativeList(items){return items.length?`<ul>${items.map(i=>`<li>${i}</li>`).join("")}</ul>`:`<div class="empty-state">当前筛选下暂无可输出内容。</div>`}
    function pick(select){return Array.from(select.selectedOptions).map(o=>o.value).filter(Boolean)}
    function fillSelect(select,options,selected){const picked=(selected||[])[0]||"";select.innerHTML="";const all=document.createElement("option");all.value="";all.textContent="全部";all.selected=!picked;select.appendChild(all);options.forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;o.selected=picked===v;select.appendChild(o);});}
    function fillSimpleSelect(select,options,selected){select.innerHTML="";options.forEach(item=>{const o=document.createElement("option");o.value=String(item.value);o.textContent=item.label;o.selected=String(item.value)===String(selected);select.appendChild(o);});}
    function matchField(value,selected){return !selected.length || selected.includes(value)}
    function currentMonth(){return DATA.meta.current_start.slice(0,7)+"-01"}
    function priorMonth(){return DATA.meta.prior_start.slice(0,7)+"-01"}
    function filterRows(rows,periodType){const start=periodType==="current"?DATA.meta.current_start:DATA.meta.prior_start;const end=periodType==="current"?DATA.meta.current_end:DATA.meta.prior_end;return rows.filter(row=>row.date>=start&&row.date<=end&&FILTER_FIELDS.every(f=>matchField(row[f],state.filters[f])));}
    function filterTT(rows,periodType){const target=periodType==="current"?currentMonth():priorMonth();return rows.filter(row=>row.month===target&&FILTER_FIELDS.every(f=>matchField(row[f],state.filters[f])));}
    function sumRows(rows,metrics){const out={};metrics.forEach(m=>out[m]=0);rows.forEach(row=>metrics.forEach(m=>out[m]+=safeNum(row[m])));return out;}
    function aggregateBy(rows,dims,metrics){const map=new Map();rows.forEach(row=>{const key=dims.map(d=>row[d]||"未归类").join("||");if(!map.has(key)){const seed={};dims.forEach(d=>seed[d]=row[d]||"未归类");metrics.forEach(m=>seed[m]=0);map.set(key,seed);}const target=map.get(key);metrics.forEach(m=>target[m]+=safeNum(row[m]));});return Array.from(map.values());}
    function indexBy(rows,key){const map=new Map();rows.forEach(row=>map.set(row[key],row));return map;}
    function formatCell(key,value){if(["gmv","innerSpend","outerSpend"].includes(key))return fmtMoney(value);if(["visitors","impressions"].includes(key))return fmtNum(value);if(["cvr","gmvMom","visitorsMom","cvrMom"].includes(key))return fmtPct(value);return value;}
    function downloadCanvas(id){const el=document.getElementById(id);if(window.Plotly&&el&&el.classList&&el.classList.contains("plot")){Plotly.downloadImage(el,{format:"png",filename:id,width:1440,height:860,scale:2});return;}const link=document.createElement("a");link.href=el.toDataURL("image/png");link.download=`${id}.png`;link.click();}
    function shortLabel(text,limit=16){const t=String(text||"");return t.length>limit?t.slice(0,limit)+"…":t;}
    function metricAxisTitle(field){return field==="cvr"?"转化率":field==="visitors"?"访客":field==="orders"?"订单量":field==="innerSpend"?"站内花费":field==="outerSpend"?"站外花费":field==="spend"?"花费":"GMV";}
    function basePlotConfig(){return {responsive:true,displaylogo:false,displayModeBar:false};}
    function basePlotFont(){return {family:"Microsoft YaHei, PingFang SC, Segoe UI, sans-serif",size:12,color:"#64748b"};}
    function renderPlot(id,data,layout,config){const el=document.getElementById(id);try{return Plotly.newPlot(el,data,layout,config);}catch(err){console.error(err);if(el){el.innerHTML=`<div class="empty-state">图表加载失败：${String(err&&err.message?err.message:err)}</div>`;}return null;}}
    function drawLine(id,series){const x=Array.from({length:series[0]?.values?.length||0},(_,i)=>i+1);renderPlot(id,series.map(s=>({type:"scatter",mode:"lines+markers",name:s.label,x,y:s.values,line:{color:s.color,width:3},marker:{size:7,color:s.color},hovertemplate:`<b>${s.label}</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>`})),{height:430,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:24,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"GMV",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawCombo(id,barsA,barsB,line){const x=Array.from({length:line.length},(_,i)=>i+1);renderPlot(id,[{type:"bar",name:"站外花费",x,y:barsA,marker:{color:"#93c5fd"},hovertemplate:"<b>站外花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"bar",name:"站内花费",x,y:barsB,marker:{color:"#dbeafe"},hovertemplate:"<b>站内花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"scatter",mode:"lines+markers",name:"店铺GMV",x,y:line,yaxis:"y2",line:{color:"#274472",width:3},marker:{size:7,color:"#274472"},hovertemplate:"<b>店铺GMV</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:58,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:"GMV",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const data=rows.slice(0,state.topN).reverse();const isPct=valueField==="cvr";renderPlot(id,[{type:"bar",orientation:"h",showlegend:false,y:data.map(r=>shortLabel(r[labelField],18)),x:data.map(r=>safeNum(r[valueField])),text:data.map(r=>formatter(r[valueField])),textposition:"outside",cliponaxis:false,marker:{color:"#8db9ff"},customdata:data.map(r=>[r[labelField],safeNum(r[deltaField])]),hovertemplate:`<b>%{customdata[0]}</b><br>${metricAxisTitle(valueField)}：${isPct?"%{x:.1%}":"%{x:,}"}<br>环比：%{customdata[1]:+.1%}<extra></extra>`}],{height:Math.max(390,120+data.length*34),paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:180,r:80,t:16,b:36},hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),xaxis:{title:metricAxisTitle(valueField),gridcolor:"#eef3fb",zeroline:false,tickformat:isPct?".0%":",.0f",tickfont:{color:"#64748b"}},yaxis:{showgrid:false,tickfont:{color:"#334155"},automargin:true}},basePlotConfig());}
    function drawPaired(id,rows){const data=rows.slice(0,state.topN);renderPlot(id,[{type:"bar",name:"Shopee",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.shopee)),marker:{color:"#60a5fa"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>Shopee销量：%{y:,}<extra></extra>"},{type:"bar",name:"TT",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.tt)),marker:{color:"#cfe1ff"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>TT销量：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:60,r:24,t:16,b:90},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"销量",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function setupMeta(){document.title=DATA.meta.title;document.getElementById("page-title").textContent=DATA.meta.title;document.getElementById("page-subtitle").textContent=DATA.meta.subtitle;document.getElementById("period-pill").textContent=`当前周期 ${DATA.meta.current_start} 至 ${DATA.meta.current_end}`;document.getElementById("compare-pill").textContent=`对比周期 ${DATA.meta.prior_start} 至 ${DATA.meta.prior_end}`;const sourceNote=document.getElementById("source-note");if(sourceNote){sourceNote.innerHTML="";sourceNote.style.display="none";}}
    function updateFilterControls(reset=false){FILTER_FIELDS.forEach(f=>fillSelect(document.getElementById(`${f}-filter`),DATA.filters[f],state.filters[f]));document.getElementById("filter-summary").textContent=FILTER_FIELDS.map(f=>`${f==="shop"?"店铺":f==="group"?"分组":f==="category"?"品类":"产品"}：${state.filters[f].length?state.filters[f].join(" / "):"全部"}`).join(" ｜ ");if(reset){FILTER_FIELDS.forEach(f=>{const select=document.getElementById(`${f}-filter`);Array.from(select.options).forEach(o=>o.selected=state.filters[f].includes(o.value));});}}
    function overallMetrics(){const cs=sumRows(filterRows(DATA.shop_daily,"current"),["gmv","visitors","orders"]);const ps=sumRows(filterRows(DATA.shop_daily,"prior"),["gmv","visitors","orders"]);const co=sumRows(filterRows(DATA.outer_daily,"current"),["outer_spend","impressions","clicks"]);const po=sumRows(filterRows(DATA.outer_daily,"prior"),["outer_spend","impressions","clicks"]);const ci=sumRows(filterRows(DATA.inner_daily,"current"),["inner_spend"]);const pi=sumRows(filterRows(DATA.inner_daily,"prior"),["inner_spend"]);const current={...cs,...co,...ci};const prior={...ps,...po,...pi};current.totalSpend=current.outer_spend+current.inner_spend;prior.totalSpend=prior.outer_spend+prior.inner_spend;current.cvr=ratio(current.orders,current.visitors);prior.cvr=ratio(prior.orders,prior.visitors);current.ctr=ratio(current.clicks,current.impressions);prior.ctr=ratio(prior.clicks,prior.impressions);current.roi=ratio(current.gmv,current.totalSpend);prior.roi=ratio(prior.gmv,prior.totalSpend);current.aov=ratio(current.gmv,current.orders);prior.aov=ratio(prior.gmv,prior.orders);return {current,prior};}
    function renderSummary(metrics){const cards=[["GMV",fmtMoney(metrics.current.gmv),pctChange(metrics.current.gmv,metrics.prior.gmv),`访客 ${fmtNum(metrics.current.visitors)}，CVR ${fmtPct(metrics.current.cvr)}`],["总花费",fmtMoney(metrics.current.totalSpend),pctChange(metrics.current.totalSpend,metrics.prior.totalSpend),`站外 ${fmtMoney(metrics.current.outer_spend)} / 站内 ${fmtMoney(metrics.current.inner_spend)}`],["站外曝光",fmtNum(metrics.current.impressions),pctChange(metrics.current.impressions,metrics.prior.impressions),`点击 ${fmtNum(metrics.current.clicks)}，CTR ${fmtPct(metrics.current.ctr)}`],["访客",fmtNum(metrics.current.visitors),pctChange(metrics.current.visitors,metrics.prior.visitors),`订单 ${fmtNum(metrics.current.orders)}，CVR ${fmtPct(metrics.current.cvr)}`],["ROI",fmtNum(metrics.current.roi,2),pctChange(metrics.current.roi,metrics.prior.roi),`AOV ${fmtMoney(metrics.current.aov)}`]];document.getElementById("summary-grid").innerHTML=cards.map(([title,value,delta,note])=>`<div class="card"><div class="section-label">${title}</div><div class="kpi-value">${value}</div><div class="delta ${deltaClass(delta)}">${fmtDelta(delta)}</div><div class="muted">${note}</div></div>`).join("");}
    function buildDimensionRows(dim){const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);return Array.from(keys).map(key=>{const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders),pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders);return {label:key,gmv:g,visitors:v,orders:o,cvr:ratio(o,v),innerSpend:safeNum(current.inner_spend),outerSpend:safeNum(current.outer_spend),impressions:safeNum(current.impressions),gmvMom:pctChange(g,pg),visitorsMom:pctChange(v,pv),cvrMom:pctChange(ratio(o,v),ratio(poo,pv))};}).filter(r=>r.label&&r.label!=="未归类").sort((a,b)=>b.gmv-a.gmv);}
    function buildModuleRows(){const current=aggregateBy(filterRows(DATA.outer_module_daily,"current"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);const prior=aggregateBy(filterRows(DATA.outer_module_daily,"prior"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);const p=indexBy(prior,"module");return current.map(r=>{const pr=p.get(r.module)||{};const roi=ratio(r.ad_gmv,r.outer_spend);return {label:r.module,spend:safeNum(r.outer_spend),impressions:safeNum(r.impressions),clicks:safeNum(r.clicks),ctr:ratio(r.clicks,r.impressions),roi,spendMom:pctChange(safeNum(r.outer_spend),safeNum(pr.outer_spend))};}).filter(r=>r.label&&r.label!=="未归类").sort((a,b)=>b.spend-a.spend);}
    function buildDiagnosisText(row,level){const spend=row.innerSpend+row.outerSpend;if(row.gmvMom>.1&&row.cvrMom>=0)return [`${level}增长健康，GMV 与转化同步改善，属于当前增长驱动。`,`继续保量，优先把预算集中到高转化链接与强素材。`];if(row.visitorsMom>.1&&row.cvrMom<-.05)return [`${level}流量放大但转化承接偏弱，更多是“点击进来但没有完成转化”。`,`优先回查 PDP 价格力、优惠机制和素材承诺一致性。`];if(spend>0&&row.gmv<spend*2)return [`${level}花费承压，当前 GMV 对花费覆盖偏弱，效率拖累明显。`,`削减低效扩量，把预算向重点套组与高转化链接回流。`];if(row.gmvMom<-.1&&row.visitorsMom<0)return [`${level}规模走弱，核心问题更偏流量不足。`,`补足曝光入口，结合素材上新与预算回补恢复基础流量盘。`];return [`${level}表现中性，暂未形成清晰增长驱动。`,`保持小步快跑测试，继续验证素材、offer 和预算结构。`];}
    function buildDiagnosisRows(kind){return buildDimensionRows(kind==="category"?"category":"product").map(row=>{const [diagnosis,todo]=buildDiagnosisText(row,kind==="category"?`${row.label}品类`:row.label);return {...row,diagnosis,todo};});}
    function buildBundleRows(){const products=buildDiagnosisRows("product");const map=indexBy(products,"label");return DATA.focus_bundles.map(name=>map.get(name)||{label:name,gmv:0,visitors:0,cvr:0,innerSpend:0,outerSpend:0,impressions:0,gmvMom:0,visitorsMom:0,cvrMom:0,diagnosis:"当前筛选范围内暂无有效店铺或投放数据。",todo:"先确认映射与素材是否已上线，再决定是否放量。"});}
function buildTTGapRows(){return filterTT(DATA.tt_monthly,"current").map(r=>({label:r.product,shopee:safeNum(r.shopee_sales),tt:safeNum(r.tt_sales),gap:safeNum(r.shopee_sales)-safeNum(r.tt_sales)})).sort((a,b)=>(safeNum(b.shopee)-safeNum(a.shopee))||(safeNum(b.tt)-safeNum(a.tt))||String(a.label).localeCompare(String(b.label),"zh-CN"));}
    function dailySeries(){const days=Array.from({length:DATA.meta.day_count},(_,i)=>i+1);const cur=indexBy(aggregateBy(filterRows(DATA.shop_daily,"current"),["day"],["gmv"]), "day");const pri=indexBy(aggregateBy(filterRows(DATA.shop_daily,"prior"),["day"],["gmv"]), "day");const out=indexBy(aggregateBy(filterRows(DATA.outer_daily,"current"),["day"],["outer_spend"]), "day");const inn=indexBy(aggregateBy(filterRows(DATA.inner_daily,"current"),["day"],["inner_spend"]), "day");return {days,gmvCurrent:days.map(d=>safeNum((cur.get(d)||{}).gmv)),gmvPrior:days.map(d=>safeNum((pri.get(d)||{}).gmv)),outerSpend:days.map(d=>safeNum((out.get(d)||{}).outer_spend)),innerSpend:days.map(d=>safeNum((inn.get(d)||{}).inner_spend))};}
    function clearCanvas(canvas){const ctx=canvas.getContext("2d");ctx.clearRect(0,0,canvas.width,canvas.height);ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);return ctx;}
    function area(canvas){return {x:58,y:24,w:canvas.width-96,h:canvas.height-72};}
    function drawLegend(ctx,items,x,y){let pos=x;items.forEach(item=>{ctx.fillStyle=item.color;ctx.fillRect(pos,y,12,12);ctx.fillStyle="#5b6673";ctx.font="12px Microsoft YaHei";ctx.fillText(item.label,pos+18,y+10);pos+=ctx.measureText(item.label).width+50;});}
    function drawLine(id,series){const canvas=document.getElementById(id);const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...series.flatMap(s=>s.values));const step=a.w/Math.max(series[0].values.length-1,1);ctx.strokeStyle="rgba(91,102,115,.18)";ctx.fillStyle="#5b6673";ctx.font="12px Microsoft YaHei";for(let i=0;i<5;i++){const y=a.y+a.h-(i/4)*a.h;ctx.beginPath();ctx.moveTo(a.x,y);ctx.lineTo(a.x+a.w,y);ctx.stroke();ctx.fillText(fmtMoney(max*i/4),10,y+4);}series.forEach(s=>{ctx.strokeStyle=s.color;ctx.lineWidth=3;ctx.beginPath();s.values.forEach((v,idx)=>{const x=a.x+step*idx;const y=a.y+a.h-(v/max)*a.h;idx?ctx.lineTo(x,y):ctx.moveTo(x,y)});ctx.stroke();s.values.forEach((v,idx)=>{const x=a.x+step*idx;const y=a.y+a.h-(v/max)*a.h;ctx.fillStyle=s.color;ctx.beginPath();ctx.arc(x,y,4,0,Math.PI*2);ctx.fill();});});ctx.textAlign="center";for(let d=1;d<=DATA.meta.day_count;d++){ctx.fillStyle="#5b6673";ctx.fillText(String(d),a.x+step*(d-1),canvas.height-28);}drawLegend(ctx,series.map(s=>({label:s.label,color:s.color})),a.x,canvas.height-18);}
    function drawCombo(id,barsA,barsB,line){const canvas=document.getElementById(id);const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...barsA,...barsB,...line);const w=a.w/Math.max(barsA.length,1);ctx.strokeStyle="rgba(91,102,115,.18)";ctx.fillStyle="#5b6673";ctx.font="12px Microsoft YaHei";for(let i=0;i<5;i++){const y=a.y+a.h-(i/4)*a.h;ctx.beginPath();ctx.moveTo(a.x,y);ctx.lineTo(a.x+a.w,y);ctx.stroke();ctx.fillText(fmtMoney(max*i/4),10,y+4);}barsA.forEach((v,idx)=>{const x=a.x+w*idx+w*.1;const bw=w*.26;const h=(v/max)*a.h;const h2=(barsB[idx]/max)*a.h;ctx.fillStyle="rgba(14,102,85,.72)";ctx.fillRect(x,a.y+a.h-h,bw,h);ctx.fillStyle="rgba(196,108,47,.72)";ctx.fillRect(x+bw+8,a.y+a.h-h2,bw,h2);});ctx.strokeStyle="#1f2732";ctx.lineWidth=3;ctx.beginPath();line.forEach((v,idx)=>{const x=a.x+w*idx+w*.62;const y=a.y+a.h-(v/max)*a.h;idx?ctx.lineTo(x,y):ctx.moveTo(x,y)});ctx.stroke();for(let d=1;d<=barsA.length;d++){ctx.textAlign="center";ctx.fillStyle="#5b6673";ctx.fillText(String(d),a.x+w*(d-1)+w*.45,canvas.height-28);}drawLegend(ctx,[{label:"站外花费",color:"rgba(14,102,85,.72)"},{label:"站内花费",color:"rgba(196,108,47,.72)"},{label:"店铺GMV",color:"#1f2732"}],a.x,canvas.height-18);}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const canvas=document.getElementById(id);const ctx=clearCanvas(canvas);const a=area(canvas);const data=rows.slice(0,8);const max=Math.max(1,...data.map(r=>safeNum(r[valueField])));const gap=a.h/Math.max(data.length,1);data.forEach((row,idx)=>{const y=a.y+idx*gap+10;const width=(safeNum(row[valueField])/max)*(a.w*.7);ctx.fillStyle="rgba(14,102,85,.72)";ctx.fillRect(a.x+150,y,width,gap*.45);ctx.fillStyle="#1f2732";ctx.font="12px Microsoft YaHei";ctx.textAlign="left";ctx.fillText(row[labelField],a.x,y+12);ctx.fillStyle="#5b6673";ctx.fillText(formatter(row[valueField]),a.x+160+width+10,y+12);if(deltaField){ctx.fillStyle=safeNum(row[deltaField])>=0?"#0e6655":"#a94442";ctx.fillText(fmtDelta(row[deltaField]),a.x+a.w-48,y+12);}});drawLegend(ctx,[{label:"当前周期",color:"rgba(14,102,85,.72)"}],a.x,canvas.height-18);}
    function drawPaired(id,rows){const canvas=document.getElementById(id);const ctx=clearCanvas(canvas);const a=area(canvas);const data=rows.slice(0,8);const max=Math.max(1,...data.flatMap(r=>[r.shopee,r.tt]));const w=a.w/Math.max(data.length,1);data.forEach((row,idx)=>{const x=a.x+w*idx;const bw=w*.24;const h1=(row.shopee/max)*a.h;const h2=(row.tt/max)*a.h;ctx.fillStyle="rgba(14,102,85,.72)";ctx.fillRect(x+w*.18,a.y+a.h-h1,bw,h1);ctx.fillStyle="rgba(196,108,47,.72)";ctx.fillRect(x+w*.18+bw+8,a.y+a.h-h2,bw,h2);ctx.fillStyle="#5b6673";ctx.font="11px Microsoft YaHei";ctx.textAlign="center";ctx.fillText(row.label.length>8?row.label.slice(0,8)+"…":row.label,x+w*.44,canvas.height-28);});drawLegend(ctx,[{label:"Shopee",color:"rgba(14,102,85,.72)"},{label:"TT",color:"rgba(196,108,47,.72)"}],a.x,canvas.height-18);}
    function renderTable(kind,rows){const controls=document.getElementById(`${kind}-table-controls`);const holder=document.getElementById(`${kind}-table`);const ts=state.tableState[kind];const columns=TABLE_COLUMNS[kind];controls.innerHTML=`<div class="table-config">${columns.map(([k,l])=>`<div><label>${l}</label><div class="toggle-chip"><input type="checkbox" data-kind="${kind}" data-column="${k}" ${ts.visible[k]?"checked":""}><input type="text" data-kind="${kind}" data-label="${k}" value="${ts.labels[k]}"></div></div>`).join("")}<div><label>行控制</label><button class="mini-btn" data-reset="${kind}">恢复全部隐藏行</button></div><div><label>导出</label><button class="mini-btn" data-export="${kind}">下载当前表格 CSV</button></div></div>`;const visible=columns.filter(([k])=>ts.visible[k]);const list=rows.filter(r=>!ts.hiddenRows.has(r.label));if(!list.length){holder.innerHTML=`<div class="empty-state">当前筛选下暂无${kind==="category"?"品类":"产品"}诊断表数据。</div>`;bindTableActions(kind,rows);return;}holder.innerHTML=`<div class="table-wrap"><table><thead><tr>${visible.map(([k])=>`<th>${ts.labels[k]}</th>`).join("")}<th>操作</th></tr></thead><tbody>${list.map(r=>`<tr>${visible.map(([k])=>`<td>${formatCell(k,r[k])}</td>`).join("")}<td><button class="mini-btn" data-hide="${kind}" data-row="${r.label}">隐藏</button></td></tr>`).join("")}</tbody></table></div>`;bindTableActions(kind,rows);}
    function bindTableActions(kind,rows){document.querySelectorAll(`input[data-kind="${kind}"][data-column]`).forEach(i=>i.onchange=()=>{state.tableState[kind].visible[i.dataset.column]=i.checked;renderAll();});document.querySelectorAll(`input[data-kind="${kind}"][data-label]`).forEach(i=>i.oninput=()=>{state.tableState[kind].labels[i.dataset.label]=i.value||TABLE_COLUMNS[kind].find(([k])=>k===i.dataset.label)[1];renderTable(kind,rows);});document.querySelectorAll(`button[data-hide="${kind}"]`).forEach(b=>b.onclick=()=>{state.tableState[kind].hiddenRows.add(b.dataset.row);renderAll();});const reset=document.querySelector(`button[data-reset="${kind}"]`);if(reset)reset.onclick=()=>{state.tableState[kind].hiddenRows.clear();renderAll();};const exp=document.querySelector(`button[data-export="${kind}"]`);if(exp)exp.onclick=()=>exportCsv(kind,rows);}
    function exportCsv(kind,rows){const cols=TABLE_COLUMNS[kind].filter(([k])=>state.tableState[kind].visible[k]);const header=cols.map(([k])=>state.tableState[kind].labels[k]).join(",");const lines=[header];rows.filter(r=>!state.tableState[kind].hiddenRows.has(r.label)).forEach(row=>{lines.push(cols.map(([k])=>`"${String(formatCell(k,row[k])).replaceAll('"','""')}"`).join(","));});const blob=new Blob(["\uFEFF"+lines.join("\\n")],{type:"text/csv;charset=utf-8;"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=`${kind}_diagnosis.csv`;a.click();URL.revokeObjectURL(url);}
    function renderNarratives(metrics,groupRows,moduleRows,categoryRows,bundleRows,ttRows){const ttSum=sumRows(filterTT(DATA.tt_monthly,"current"),["tt_sales","shopee_sales"]);const gap=ttSum.shopee_sales-ttSum.tt_sales;document.getElementById("overall-narrative").innerHTML=`<p>当前筛选范围内，店铺 GMV 为 ${fmtMoney(metrics.current.gmv)}，较 3 月同期 ${fmtDelta(pctChange(metrics.current.gmv,metrics.prior.gmv))}；访客 ${fmtNum(metrics.current.visitors)}，CVR ${fmtPct(metrics.current.cvr)}。</p><p>站外花费 ${fmtMoney(metrics.current.outer_spend)}、站内花费 ${fmtMoney(metrics.current.inner_spend)}，当前 ROI 为 ${fmtNum(metrics.current.roi,2)}。Shopee 销量 ${fmtNum(ttSum.shopee_sales)}，TT 销量 ${fmtNum(ttSum.tt_sales)}，GAP 为 ${gap>=0?"领先":"落后"} ${fmtNum(Math.abs(gap))}。</p><p><strong>下一步：</strong></p>${narrativeList(["若 GMV 增速低于花费增速，优先拆分流量扩张与转化承接问题。","4 月继续把 2 件套作为重点资源方向，避免只做单品放量。"])}`;
      document.getElementById("spend-narrative").innerHTML=`<p>${pctChange(metrics.current.gmv,metrics.prior.gmv)>=pctChange(metrics.current.totalSpend,metrics.prior.totalSpend)?"本期 GMV 增速高于或接近花费增速，放量尚未明显稀释产出。":"本期花费增速高于 GMV 增速，效率承压，需要回到产品承接和素材命中率上做修正。"}</p><p>当前站外花费占总花费 ${fmtPct(ratio(metrics.current.outer_spend,metrics.current.totalSpend))}，站内花费占比 ${fmtPct(ratio(metrics.current.inner_spend,metrics.current.totalSpend))}。</p><p><strong>下一步：</strong></p>${narrativeList(["按日复盘站外花费与 GMV 的错位天，定位是否为素材衰退或站内承接波动。","对站内广告做产品级 ROI 审视，避免继续抬高低转化商品流量成本。"])}`;
      if(groupRows.length){const top=groupRows[0],weak=[...groupRows].sort((a,b)=>a.gmvMom-b.gmvMom)[0];document.getElementById("group-narrative").innerHTML=`<p>${top.label} 是当前 GMV 第一分组，GMV ${fmtMoney(top.gmv)}，环比上月同期 ${fmtDelta(top.gmvMom)}。</p><p>${weak.label} 是当前增速偏弱分组，GMV 环比 ${fmtDelta(weak.gmvMom)}，需要重点检查是否有流量放大但承接不足的问题。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续保量，同时追加高转化素材。`,`${weak.label} 优先做素材命中率和店铺承接排查。`])}`;}else{document.getElementById("group-narrative").innerHTML='<div class="empty-state">当前筛选下暂无分组数据。</div>';}
      if(moduleRows.length){const top=moduleRows[0],best=[...moduleRows].sort((a,b)=>b.roi-a.roi)[0],weak=[...moduleRows].sort((a,b)=>a.roi-b.roi)[0];document.getElementById("module-narrative").innerHTML=`<p>${top.label} 是当前预算最高模块，站外花费 ${fmtMoney(top.spend)}，CTR ${fmtPct(top.ctr)}。</p><p>${best.label} 的站外 ROI 最高，为 ${fmtNum(best.roi,2)}；${weak.label} 的 ROI 偏低，更像当前效率拖累点。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 若继续承担主要预算，应同步核查 CTR 和素材疲劳。`,`${weak.label} 优先收缩低效受众或替换素材方向。`])}`;}else{document.getElementById("module-narrative").innerHTML='<div class="empty-state">当前筛选下暂无模块数据。</div>';}
      if(categoryRows.length){const top=categoryRows[0],best=[...categoryRows].sort((a,b)=>b.cvr-a.cvr)[0],weak=[...categoryRows].sort((a,b)=>a.gmvMom-b.gmvMom)[0];document.getElementById("category-narrative").innerHTML=`<p>${top.label} 是当前 GMV 最大品类，GMV ${fmtMoney(top.gmv)}，访客 ${fmtNum(top.visitors)}。</p><p>${best.label} 的 CVR 最好，为 ${fmtPct(best.cvr)}；${weak.label} 的 GMV 环比 ${fmtDelta(weak.gmvMom)}，需要警惕规模下滑。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续稳住流量入口，并把预算向主销 SKU 与套组集中。`,`${weak.label} 若继续投放，必须先修正转化承接或价格钩子。`])}`;}else{document.getElementById("category-narrative").innerHTML='<div class="empty-state">当前筛选下暂无品类数据。</div>';}
      if(ttRows.length){const best=[...ttRows].sort((a,b)=>b.gap-a.gap)[0],worst=[...ttRows].sort((a,b)=>a.gap-b.gap)[0];document.getElementById("tt-gap-narrative").innerHTML=`<p>${best.label} 当前 Shopee 相对 TT 领先 ${fmtNum(Math.max(best.gap,0))}；${worst.label} 相对 TT ${worst.gap<0?"落后":"领先"} ${fmtNum(Math.abs(worst.gap))}。</p><p>当前版位用于识别需要补销量的产品，而不是单看投放成本。</p><p><strong>下一步：</strong></p>${narrativeList([worst.gap<0?`${worst.label} 是当前最需要补差的产品，建议结合素材和店铺资源做专项修复。`:"当前筛选下暂无明显 TT 销量缺口，可继续稳态放量。","TT GAP 需与 2 件套策略联动，优先用 377 combo 与新组合测试寻找补量点。"])}`;}else{document.getElementById("tt-gap-narrative").innerHTML='<div class="empty-state">当前筛选下暂无 TT 对比数据。</div>';}
      const focus377=bundleRows.find(r=>r.label==="2件套-377精华+面霜");const weakBundle=bundleRows.find(r=>r.label==="2件套-精华+面霜");document.getElementById("highlight-box").innerHTML=narrativeList([groupRows[0]?`${groupRows[0].label} 是当前最大 GMV 驱动分组，贡献 ${fmtMoney(groupRows[0].gmv)}。`:"",categoryRows[0]?`${categoryRows[0].label} 保持当前最大品类体量。`:"",focus377&&focus377.outerSpend>0?`377 combo 仍在持续投放，当前站外花费 ${fmtMoney(focus377.outerSpend)}，仍是 2 件套核心规模盘。`:""].filter(Boolean));
      document.getElementById("weakness-box").innerHTML=narrativeList([weakBundle&&weakBundle.outerSpend<=0?`2件套-精华+面霜 当前几乎没有外投承接，符合“素材尚未产出”的现状，4 月新增量存在明显缺口。`:"",bundleRows.filter(r=>r.label!=="2件套-377精华+面霜").sort((a,b)=>a.outerSpend-b.outerSpend)[0]?`${bundleRows.filter(r=>r.label!=="2件套-377精华+面霜").sort((a,b)=>a.outerSpend-b.outerSpend)[0].label} 当前外投量级偏小，尚未形成有效验证样本。`:"",ttRows.find(r=>r.gap<0)?`${ttRows.find(r=>r.gap<0).label} 的 Shopee 销量低于 TT ${fmtNum(Math.abs(ttRows.find(r=>r.gap<0).gap))}，需要在 4 月补差。`:""].filter(Boolean));
      document.getElementById("bundle-box").innerHTML=`<p>${bundleRows.map(r=>`${r.label}：站外花费 ${fmtMoney(r.outerSpend)}，店铺 GMV ${fmtMoney(r.gmv)}，CVR ${fmtPct(r.cvr)}。${r.label==="2件套-377精华+面霜"?"377 combo 持续投放，仍是 2 件套基础规模盘。":r.label==="2件套-精华+面霜"?"该套组对应 PDRN 面霜 + 视黄醇精华，当前更大的问题不是转化，而是素材未上线导致验证样本不足。":r.label==="PDRN+素颜霜2件套"?"该套组已有少量投放，建议以妆前/素颜场景切入测试。":"该套组目前投放量不高，更适合先验证内容命题与连带需求。"}`).join("</p><p>")}</p><p><strong>下一步：</strong></p>${narrativeList(["2件套-377精华+面霜 继续做规模盘，同时监控 ROI 和访客转化。","2件套-精华+面霜 尽快补齐素材并建立最小验证预算。","PDRN+素颜霜2件套 与 2件套-磨皮棉片+美白面膜 先以小预算多素材快速验证。"])}`;
      document.getElementById("action-box").innerHTML=narrativeList([groupRows[0]?`优先给 ${groupRows[0].label} 分组保量，因为它仍是当前最大的 GMV 驱动盘。`:"当前筛选无有效分组数据，请先放宽筛选。",buildDiagnosisRows("product").filter(r=>r.outerSpend+r.innerSpend>0).sort((a,b)=>(a.gmv/Math.max(a.outerSpend+a.innerSpend,1))-(b.gmv/Math.max(b.outerSpend+b.innerSpend,1)))[0]?`收缩 ${buildDiagnosisRows("product").filter(r=>r.outerSpend+r.innerSpend>0).sort((a,b)=>(a.gmv/Math.max(a.outerSpend+a.innerSpend,1))-(b.gmv/Math.max(b.outerSpend+b.innerSpend,1)))[0].label} 的低效花费，把预算向高转化链接和重点 2 件套回流。`:"当前筛选无足够产品效率样本。",ttRows.find(r=>r.gap<0)?`针对 ${ttRows.find(r=>r.gap<0).label} 制定 TT GAP 补差动作，优先看是否需要额外素材、价格钩子或站内流量扶持。`:"当前筛选下暂无明显 TT 销量缺口，可继续稳态放量。"]);
      document.getElementById("data-note-box").innerHTML=narrativeList([`当前报告固定真实业务数据源为项目 Google Sheet 主表，链接已写死到项目配置中。`,`4 月复盘按 ${DATA.meta.current_start} 至 ${DATA.meta.current_end} 的月累计计算，并与 ${DATA.meta.prior_start} 至 ${DATA.meta.prior_end} 做同天数比较。`,`GMV、访客、订单、CVR 均以店铺 GMV 表口径为准；CVR = 订单数 / 链接访客。`,`站外广告花费按 USD × 7 换算人民币后进入报表。`,`模块分析使用站外广告维度，因为店铺 GMV 数据不带模块字段。`,`重点产品映射统一回站内命名，PDRN面霜+视黄醇精华按 2件套-精华+面霜 归并。`]);}
    function bindGlobal(){FILTER_FIELDS.forEach(f=>document.getElementById(`${f}-filter`).addEventListener("change",()=>{state.filters[f]=pick(document.getElementById(`${f}-filter`));renderAll();}));document.getElementById("reset-filter").onclick=()=>{FILTER_FIELDS.forEach(f=>state.filters[f]=[]);renderAll(true);};document.getElementById("focus-bundle").onclick=()=>{state.filters.product=[...DATA.focus_bundles];renderAll(true);};document.querySelectorAll("[data-download]").forEach(b=>b.onclick=()=>downloadCanvas(b.dataset.download));}
    function renderAll(reset=false){updateFilterControls(reset);const metrics=overallMetrics();renderSummary(metrics);const daily=dailySeries();drawLine("gmv-trend",[{label:"4月当前",color:"#0e6655",values:daily.gmvCurrent},{label:"3月同期",color:"#c46c2f",values:daily.gmvPrior}]);drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");drawBars("group-chart",groupRows,"label","gmv","gmvMom",fmtMoney);drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);drawBars("category-chart",categoryRows,"label","gmv","gmvMom",fmtMoney);drawPaired("tt-gap-chart",ttRows);renderNarratives(metrics,groupRows,moduleRows,categoryRows,bundleRows,ttRows);renderTable("category",categoryDiagnosis);renderTable("product",productDiagnosis);}
    function renderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows){
      const ttSum=sumRows(filterTT(DATA.tt_monthly,"current"),["tt_sales","shopee_sales"]);
      const gap=ttSum.shopee_sales-ttSum.tt_sales;
      document.getElementById("overall-narrative").innerHTML=`<p>当前筛选范围内，店铺 GMV 为 ${fmtMoney(metrics.current.gmv)}，较 3 月同期 ${fmtDelta(pctChange(metrics.current.gmv,metrics.prior.gmv))}；访客 ${fmtNum(metrics.current.visitors)}，CVR ${fmtPct(metrics.current.cvr)}。</p><p>站外花费 ${fmtMoney(metrics.current.outer_spend)}、站内花费 ${fmtMoney(metrics.current.inner_spend)}，当前 ROI 为 ${fmtNum(metrics.current.roi,2)}。Shopee 销量 ${fmtNum(ttSum.shopee_sales)}，TT 销量 ${fmtNum(ttSum.tt_sales)}，GAP 为 ${gap>=0?"领先":"落后"} ${fmtNum(Math.abs(gap))}。</p><p><strong>下一步：</strong></p>${narrativeList(["若 GMV 增速低于花费增速，优先拆分流量扩张与转化承接问题。","4 月继续把 2 件套作为重点资源方向，避免只做单品放量。"])}`;
      document.getElementById("spend-narrative").innerHTML=`<p>${pctChange(metrics.current.gmv,metrics.prior.gmv)>=pctChange(metrics.current.totalSpend,metrics.prior.totalSpend)?"本期 GMV 增速高于或接近花费增速，放量尚未明显稀释产出。":"本期花费增速高于 GMV 增速，效率承压，需要回到产品承接和素材命中率上做修正。"}</p><p>当前站外花费占总花费 ${fmtPct(ratio(metrics.current.outer_spend,metrics.current.totalSpend))}，站内花费占比 ${fmtPct(ratio(metrics.current.inner_spend,metrics.current.totalSpend))}。</p><p><strong>下一步：</strong></p>${narrativeList(["按日复盘站外花费与 GMV 的错位天，定位是否为素材衰退或站内承接波动。","对站内广告做产品级 ROI 审视，避免继续抬高低转化商品流量成本。"])}`;
      if(groupRows.length){const top=groupRows[0],weak=[...groupRows].sort((a,b)=>a.gmvMom-b.gmvMom)[0];document.getElementById("group-narrative").innerHTML=`<p>${top.label} 是当前 GMV 第一分组，GMV ${fmtMoney(top.gmv)}，环比上月同期 ${fmtDelta(top.gmvMom)}。</p><p>${weak.label} 是当前增速偏弱分组，GMV 环比 ${fmtDelta(weak.gmvMom)}，需要重点检查是否有流量放大但承接不足的问题。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续保量，同时追加高转化素材。`,`${weak.label} 优先做素材命中率和店铺承接排查。`])}`;}else{document.getElementById("group-narrative").innerHTML='<div class="empty-state">当前筛选下暂无分组数据。</div>';}
      if(moduleRows.length){const top=moduleRows[0],best=[...moduleRows].sort((a,b)=>b.roi-a.roi)[0],weak=[...moduleRows].sort((a,b)=>a.roi-b.roi)[0];document.getElementById("module-narrative").innerHTML=`<p>${top.label} 是当前预算最高模块，站外花费 ${fmtMoney(top.spend)}，CTR ${fmtPct(top.ctr)}。</p><p>${best.label} 的站外 ROI 最高，为 ${fmtNum(best.roi,2)}；${weak.label} 的 ROI 偏低，更像当前效率拖累点。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 若继续承担主要预算，应同步核查 CTR 和素材疲劳。`,`${weak.label} 优先收缩低效受众或替换素材方向。`])}`;}else{document.getElementById("module-narrative").innerHTML='<div class="empty-state">当前筛选下暂无模块数据。</div>';}
      if(categoryRows.length){const top=categoryRows[0],best=[...categoryRows].sort((a,b)=>b.cvr-a.cvr)[0],weak=[...categoryRows].sort((a,b)=>a.gmvMom-b.gmvMom)[0];document.getElementById("category-narrative").innerHTML=`<p>${top.label} 是当前 GMV 最大品类，GMV ${fmtMoney(top.gmv)}，访客 ${fmtNum(top.visitors)}。</p><p>${best.label} 的 CVR 最好，为 ${fmtPct(best.cvr)}；${weak.label} 的 GMV 环比 ${fmtDelta(weak.gmvMom)}，需要警惕规模下滑。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续稳住流量入口，并把预算向主销 SKU 与套组集中。`,`${weak.label} 若继续投放，必须先修正转化承接或价格钩子。`])}`;}else{document.getElementById("category-narrative").innerHTML='<div class="empty-state">当前筛选下暂无品类数据。</div>';}
      if(productRows.length){const sample=productRows.filter(r=>r.gmv>=100000||r.outerSpend+r.innerSpend>=2000);const growthPool=sample.length?sample:productRows;const top=productRows[0],growth=[...growthPool].sort((a,b)=>b.gmvMom-a.gmvMom)[0],weak=[...growthPool.filter(r=>r.outerSpend+r.innerSpend>0)].sort((a,b)=>(a.gmv/Math.max(a.outerSpend+a.innerSpend,1))-(b.gmv/Math.max(b.outerSpend+b.innerSpend,1)))[0]||growthPool[0];document.getElementById("product-narrative").innerHTML=`<p>${top.label} 是当前筛选下 GMV 最高产品，GMV ${fmtMoney(top.gmv)}，访客 ${fmtNum(top.visitors)}，CVR ${fmtPct(top.cvr)}。</p><p>${growth.label} 的 GMV 环比提升 ${fmtDelta(growth.gmvMom)}，更像当前的增长弹性产品；${weak.label} 的产出效率偏弱，当前更接近预算拖累点。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续稳住链接承接和评论资产，避免规模产品因为素材疲劳掉量。`,`${growth.label} 继续验证可复制的素材角度和价格钩子，判断能否承接更多站外预算。`,`${weak.label} 先收缩低效花费，再决定是否保留站外放量。`])}`;}else{document.getElementById("product-narrative").innerHTML='<div class="empty-state">当前筛选下暂无产品分析数据。</div>';}
      if(ttRows.length){const best=[...ttRows].sort((a,b)=>b.gap-a.gap)[0],worst=[...ttRows].sort((a,b)=>a.gap-b.gap)[0];document.getElementById("tt-gap-narrative").innerHTML=`<p>${best.label} 当前 Shopee 相对 TT 领先 ${fmtNum(Math.max(best.gap,0))}；${worst.label} 相对 TT ${worst.gap<0?"落后":"领先"} ${fmtNum(Math.abs(worst.gap))}。</p><p>这个板块重点用来识别要补销量的产品，不建议只看投放成本就做判断。</p><p><strong>下一步：</strong></p>${narrativeList([worst.gap<0?`${worst.label} 是当前最需要补差的产品，建议结合素材和店铺资源做专项修复。`:"当前筛选下暂无明显 TT 销量缺口，可继续稳态放量。","TT GAP 需要与 2 件套策略联动，优先用 377 combo 与新组合测试寻找补量点。"])}`;}else{document.getElementById("tt-gap-narrative").innerHTML='<div class="empty-state">当前筛选下暂无 TT 对比数据。</div>';}
      const focus377=bundleRows.find(r=>r.label==="2件套-377精华+面霜");
      const retinolBundle=bundleRows.find(r=>r.label==="2件套-精华+面霜");
      const toneupBundle=bundleRows.find(r=>r.label==="PDRN+素颜霜2件套");
      const maskBundle=bundleRows.find(r=>r.label==="2件套-磨皮棉片+美白面膜");
      document.getElementById("highlight-box").innerHTML=narrativeList([groupRows[0]?`${groupRows[0].label} 是当前最大 GMV 驱动分组，贡献 ${fmtMoney(groupRows[0].gmv)}。`:"",productRows[0]?`${productRows[0].label} 是当前 GMV 最大单品/链接，适合作为流量承接盘持续保量。`:"",focus377&&focus377.outerSpend>0?`377 combo 仍在持续投放，当前站外花费 ${fmtMoney(focus377.outerSpend)}，仍是 2 件套核心规模盘。`:"",toneupBundle&&toneupBundle.gmvMom>0?`PDRN+素颜霜 2 件套当前呈增长态势，说明少量投放已经能撬动新增需求验证。`:""].filter(Boolean));
      document.getElementById("weakness-box").innerHTML=narrativeList([retinolBundle&&retinolBundle.outerSpend<=300?`2件套-精华+面霜 当前几乎没有外投承接，符合“素材尚未产出”的现状，4 月新增量存在明显缺口。`:"",maskBundle&&maskBundle.outerSpend<500?`2件套-磨皮棉片+美白面膜 当前外投量级仍小，样本尚不足以判断是否能成为稳定新套组。`:"",ttRows.find(r=>r.gap<0)?`${ttRows.find(r=>r.gap<0).label} 的 Shopee 销量低于 TT ${fmtNum(Math.abs(ttRows.find(r=>r.gap<0).gap))}，需要在 4 月补差。`:""].filter(Boolean));
      const bundleParagraphs=bundleRows.map(r=>{let detail="";if(r.label==="2件套-377精华+面霜"){detail="377 combo 仍是持续投放的重点 2 件套，更多承担规模盘角色。";}else if(r.label==="2件套-精华+面霜"){detail="该套组对应站内命名“2件套-精华+面霜”，即 PDRN 面霜 + 视黄醇精华；当前更大的问题不是站内承接，而是站外素材样本不足。";}else if(r.label==="PDRN+素颜霜2件套"){detail="该套组已有少量投放，且当前增长明显，更适合继续用妆前/素颜场景切题测试。";}else{detail="该套组已有少量验证，但当前体量仍小，建议继续围绕美白+磨皮联用诉求补素材。";}return `${r.label}：店铺 GMV ${fmtMoney(r.gmv)}，访客 ${fmtNum(r.visitors)}，CVR ${fmtPct(r.cvr)}，站外花费 ${fmtMoney(r.outerSpend)}。${detail}`;}).join("</p><p>");
      document.getElementById("bundle-narrative").innerHTML=`<p>${bundleParagraphs}</p><p><strong>下一步：</strong></p>${narrativeList(["2件套-377精华+面霜 继续承担规模盘，同时严控 ROI、CTR 和 CVR，不要只看放量。","2件套-精华+面霜 尽快补齐素材，并给到最小测试预算验证 PDRN 面霜 + 视黄醇精华的 Meta 端点击质量。","PDRN+素颜霜2件套 与 2件套-磨皮棉片+美白面膜 先用小预算、多素材、多场景快速试错，再决定谁进入下一轮放量。"])}`;
      document.getElementById("action-box").innerHTML=narrativeList([groupRows[0]?`优先给 ${groupRows[0].label} 分组保量，因为它仍是当前最大的 GMV 驱动盘。`:"当前筛选无有效分组数据，请先放宽筛选。",productRows.filter(r=>r.outerSpend+r.innerSpend>0).sort((a,b)=>(a.gmv/Math.max(a.outerSpend+a.innerSpend,1))-(b.gmv/Math.max(b.outerSpend+b.innerSpend,1)))[0]?`收缩 ${productRows.filter(r=>r.outerSpend+r.innerSpend>0).sort((a,b)=>(a.gmv/Math.max(a.outerSpend+a.innerSpend,1))-(b.gmv/Math.max(b.outerSpend+b.innerSpend,1)))[0].label} 的低效花费，把预算向高转化链接和重点 2 件套回流。`:"当前筛选无足够产品效率样本。",retinolBundle?`给 2件套-精华+面霜 单独建素材测试池，并把站外命名统一映射回站内口径，方便持续追踪。`:"当前筛选未覆盖 PDRN+视黄醇精华套组。",ttRows.find(r=>r.gap<0)?`针对 ${ttRows.find(r=>r.gap<0).label} 制定 TT GAP 补差动作，优先看是否需要额外素材、价格钩子或站内流量扶持。`:"当前筛选下暂无明显 TT 销量缺口，可继续稳态放量。"]);
      document.getElementById("data-note-box").innerHTML=narrativeList([`当前报告固定真实业务数据源为项目 Google Sheet 主表，链接已写死到项目配置中。`,`4 月复盘按 ${DATA.meta.current_start} 至 ${DATA.meta.current_end} 的月累计计算，并与 ${DATA.meta.prior_start} 至 ${DATA.meta.prior_end} 做同天数比较。`,`GMV、访客、订单、CVR 均以店铺 GMV 表口径为准；CVR = 订单数 / 链接访客。`,`站外广告花费按 USD × 7 换算人民币后进入报表。`,`模块分析使用站外广告维度，因为店铺 GMV 数据不带模块字段。`,`重点产品映射统一回站内命名，PDRN面霜+视黄醇精华按 2件套-精华+面霜 归并。`]);
    }
    function renderAll(reset=false){updateFilterControls(reset);const metrics=overallMetrics();renderSummary(metrics);const daily=dailySeries();drawLine("gmv-trend",[{label:"4月当前",color:"#0e6655",values:daily.gmvCurrent},{label:"3月同期",color:"#c46c2f",values:daily.gmvPrior}]);drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");drawBars("group-chart",groupRows,"label","gmv","gmvMom",fmtMoney);drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);drawBars("category-chart",categoryRows,"label","gmv","gmvMom",fmtMoney);drawBars("product-chart",productRows,"label","gmv","gmvMom",fmtMoney);drawPaired("tt-gap-chart",ttRows);drawBars("bundle-chart",bundleRows,"label","gmv","gmvMom",fmtMoney);renderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);renderTable("category",categoryDiagnosis);renderTable("product",productDiagnosis);}
    function getMetricConfig(){return {gmv:{field:"gmv",delta:"gmvMom",formatter:fmtMoney},orders:{field:"orders",delta:"ordersMom",formatter:v=>fmtNum(v)},visitors:{field:"visitors",delta:"visitorsMom",formatter:v=>fmtNum(v)},cvr:{field:"cvr",delta:"cvrMom",formatter:fmtPct},innerSpend:{field:"innerSpend",delta:"innerSpendMom",formatter:fmtMoney},outerSpend:{field:"outerSpend",delta:"outerSpendMom",formatter:fmtMoney}}[state.metricFocus]||{field:"gmv",delta:"gmvMom",formatter:fmtMoney};}
    function buildDimensionRows(dim){const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);const rows=Array.from(keys).map(key=>{const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders),pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders),cis=safeNum(current.inner_spend),pis=safeNum(prior.inner_spend),cos=safeNum(current.outer_spend),pos=safeNum(prior.outer_spend);return {label:key,gmv:g,visitors:v,orders:o,cvr:ratio(o,v),innerSpend:cis,outerSpend:cos,impressions:safeNum(current.impressions),gmvMom:pctChange(g,pg),ordersMom:pctChange(o,poo),visitorsMom:pctChange(v,pv),cvrMom:pctChange(ratio(o,v),ratio(poo,pv)),innerSpendMom:pctChange(cis,pis),outerSpendMom:pctChange(cos,pos)};}).filter(r=>r.label&&r.label!=="鏈綊绫?");const cfg=getMetricConfig();return rows.sort((a,b)=>safeNum(b[cfg.field])-safeNum(a[cfg.field]));}
    function updateFilterControls(reset=false){FILTER_FIELDS.forEach(f=>fillSelect(document.getElementById(`${f}-filter`),DATA.filters[f],state.filters[f]));fillSimpleSelect(document.getElementById("metric-filter"),METRIC_OPTIONS,state.metricFocus);fillSimpleSelect(document.getElementById("topn-filter"),TOP_N_OPTIONS.map(v=>({value:v,label:String(v)})),state.topN);document.getElementById("filter-summary").textContent=[`模块视角：${state.filters.group[0]||"全部"}`,`品类：${state.filters.category[0]||"全部"}`,`产品：${state.filters.product[0]||"全部"}`,`店铺：${state.filters.shop[0]||"全部"}`,`Top N：${state.topN}`].join(" ｜ ");if(reset){FILTER_FIELDS.forEach(f=>{const select=document.getElementById(`${f}-filter`);Array.from(select.options).forEach(o=>o.selected=(state.filters[f][0]||"")===o.value);});}}
    function renderSummary(metrics){const cards=[["店铺GMV",fmtMoney(metrics.current.gmv),pctChange(metrics.current.gmv,metrics.prior.gmv),`对比 ${DATA.meta.prior_start} ~ ${DATA.meta.prior_end}`],["订单量",fmtNum(metrics.current.orders),pctChange(metrics.current.orders,metrics.prior.orders),`AOV ${fmtMoney(metrics.current.aov)}`],["访客",fmtNum(metrics.current.visitors),pctChange(metrics.current.visitors,metrics.prior.visitors),`CTR ${fmtPct(metrics.current.ctr)}`],["转化率",fmtPct(metrics.current.cvr),pctChange(metrics.current.cvr,metrics.prior.cvr),`ROI ${fmtNum(metrics.current.roi,2)}`],["站内花费",fmtMoney(metrics.current.inner_spend),pctChange(metrics.current.inner_spend,metrics.prior.inner_spend),`占总花费 ${fmtPct(ratio(metrics.current.inner_spend,metrics.current.totalSpend))}`],["站外花费",fmtMoney(metrics.current.outer_spend),pctChange(metrics.current.outer_spend,metrics.prior.outer_spend),`站外曝光 ${fmtNum(metrics.current.impressions)}`]];document.getElementById("summary-grid").innerHTML=cards.map(([title,value,delta,note])=>`<div class="card summary-card"><div class="section-label">${title}</div><div class="kpi-value">${value}</div><span class="delta ${deltaClass(delta)}">${fmtDelta(delta)}</span><span class="muted">${note}</span></div>`).join("");}
    function drawLine(id,series){const canvas=document.getElementById(id);canvas.height=360;const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...series.flatMap(s=>s.values));const step=a.w/Math.max(series[0].values.length-1,1);ctx.strokeStyle="rgba(111,132,163,.16)";ctx.fillStyle="#6f84a3";ctx.font="12px Microsoft YaHei";for(let i=0;i<5;i++){const y=a.y+a.h-(i/4)*a.h;ctx.beginPath();ctx.moveTo(a.x,y);ctx.lineTo(a.x+a.w,y);ctx.stroke();ctx.fillText(fmtMoney(max*i/4),10,y+4);}series.forEach(s=>{ctx.strokeStyle=s.color;ctx.lineWidth=3;ctx.beginPath();s.values.forEach((v,idx)=>{const x=a.x+step*idx;const y=a.y+a.h-(v/max)*a.h;idx?ctx.lineTo(x,y):ctx.moveTo(x,y)});ctx.stroke();s.values.forEach((v,idx)=>{const x=a.x+step*idx;const y=a.y+a.h-(v/max)*a.h;ctx.fillStyle=s.color;ctx.beginPath();ctx.arc(x,y,4,0,Math.PI*2);ctx.fill();});});ctx.textAlign="center";for(let d=1;d<=DATA.meta.day_count;d++){ctx.fillStyle="#6f84a3";ctx.fillText(String(d),a.x+step*(d-1),canvas.height-28);}drawLegend(ctx,series.map(s=>({label:s.label,color:s.color})),a.x,canvas.height-18);}
    function drawCombo(id,barsA,barsB,line){const canvas=document.getElementById(id);canvas.height=360;const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...barsA,...barsB,...line);const w=a.w/Math.max(barsA.length,1);ctx.strokeStyle="rgba(111,132,163,.16)";ctx.fillStyle="#6f84a3";ctx.font="12px Microsoft YaHei";for(let i=0;i<5;i++){const y=a.y+a.h-(i/4)*a.h;ctx.beginPath();ctx.moveTo(a.x,y);ctx.lineTo(a.x+a.w,y);ctx.stroke();ctx.fillText(fmtMoney(max*i/4),10,y+4);}barsA.forEach((v,idx)=>{const x=a.x+w*idx+w*.1;const bw=w*.26;const h=(v/max)*a.h;const h2=(barsB[idx]/max)*a.h;ctx.fillStyle="rgba(47,107,255,.72)";ctx.fillRect(x,a.y+a.h-h,bw,h);ctx.fillStyle="rgba(132,181,255,.78)";ctx.fillRect(x+bw+8,a.y+a.h-h2,bw,h2);});ctx.strokeStyle="#1e3557";ctx.lineWidth=3;ctx.beginPath();line.forEach((v,idx)=>{const x=a.x+w*idx+w*.62;const y=a.y+a.h-(v/max)*a.h;idx?ctx.lineTo(x,y):ctx.moveTo(x,y)});ctx.stroke();for(let d=1;d<=barsA.length;d++){ctx.textAlign="center";ctx.fillStyle="#6f84a3";ctx.fillText(String(d),a.x+w*(d-1)+w*.45,canvas.height-28);}drawLegend(ctx,[{label:"站外花费",color:"rgba(47,107,255,.72)"},{label:"站内花费",color:"rgba(132,181,255,.78)"},{label:"店铺GMV",color:"#1e3557"}],a.x,canvas.height-18);}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const data=rows.slice(0,state.topN);const canvas=document.getElementById(id);canvas.height=Math.max(360,96+data.length*28);const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...data.map(r=>safeNum(r[valueField])));const gap=a.h/Math.max(data.length,1);data.forEach((row,idx)=>{const y=a.y+idx*gap+8;const width=(safeNum(row[valueField])/max)*(a.w*.68);ctx.fillStyle="rgba(47,107,255,.72)";ctx.fillRect(a.x+150,y,width,Math.max(12,gap*.48));ctx.fillStyle="#1e3557";ctx.font="12px Microsoft YaHei";ctx.textAlign="left";ctx.fillText(row[labelField],a.x,y+12);ctx.fillStyle="#6f84a3";ctx.fillText(formatter(row[valueField]),a.x+160+width+10,y+12);if(deltaField){ctx.fillStyle=safeNum(row[deltaField])>=0?"#15a05d":"#d96464";ctx.fillText(fmtDelta(row[deltaField]),a.x+a.w-52,y+12);}});drawLegend(ctx,[{label:"当前周期",color:"rgba(47,107,255,.72)"}],a.x,canvas.height-18);}
    function drawPaired(id,rows){const data=rows.slice(0,state.topN);const canvas=document.getElementById(id);canvas.height=Math.max(360,100+Math.ceil(data.length/Math.max(1,Math.floor(canvas.width/120)))*40);const ctx=clearCanvas(canvas);const a=area(canvas);const max=Math.max(1,...data.flatMap(r=>[r.shopee,r.tt]));const w=a.w/Math.max(data.length,1);data.forEach((row,idx)=>{const x=a.x+w*idx;const bw=Math.min(22,w*.24);const h1=(row.shopee/max)*a.h;const h2=(row.tt/max)*a.h;ctx.fillStyle="rgba(47,107,255,.72)";ctx.fillRect(x+w*.18,a.y+a.h-h1,bw,h1);ctx.fillStyle="rgba(132,181,255,.78)";ctx.fillRect(x+w*.18+bw+8,a.y+a.h-h2,bw,h2);ctx.fillStyle="#6f84a3";ctx.font="11px Microsoft YaHei";ctx.textAlign="center";ctx.fillText(row.label.length>8?row.label.slice(0,8)+"…":row.label,x+w*.44,canvas.height-28);});drawLegend(ctx,[{label:"Shopee",color:"rgba(47,107,255,.72)"},{label:"TT",color:"rgba(132,181,255,.78)"}],a.x,canvas.height-18);}
    function setupMeta(){document.title=DATA.meta.title;document.getElementById("page-title").textContent=DATA.meta.title;document.getElementById("page-subtitle").textContent=DATA.meta.subtitle;document.getElementById("period-pill").textContent=`当前周期 ${DATA.meta.current_start} 至 ${DATA.meta.current_end}`;document.getElementById("compare-pill").textContent=`对比周期 ${DATA.meta.prior_start} 至 ${DATA.meta.prior_end}`;const sourceNote=document.getElementById("source-note");if(sourceNote){sourceNote.innerHTML="";sourceNote.style.display="none";}}
    function shortLabel(text,limit=16){const t=String(text||"");return t.length>limit?t.slice(0,limit)+"…":t;}
    function metricAxisTitle(field){return field==="cvr"?"转化率":field==="visitors"?"访客":field==="orders"?"订单量":field==="innerSpend"?"站内花费":field==="outerSpend"?"站外花费":field==="spend"?"花费":"GMV";}
    function basePlotConfig(){return {responsive:true,displaylogo:false,displayModeBar:false};}
    function basePlotFont(){return {family:"Microsoft YaHei, PingFang SC, Segoe UI, sans-serif",size:12,color:"#64748b"};}
    function renderPlot(id,data,layout,config){const el=document.getElementById(id);try{return Plotly.newPlot(el,data,layout,config);}catch(err){console.error(err);if(el){el.innerHTML=`<div class="empty-state">图表加载失败：${String(err&&err.message?err.message:err)}</div>`;}return null;}}
    function drawLine(id,series){const x=Array.from({length:series[0]?.values?.length||0},(_,i)=>i+1);renderPlot(id,series.map(s=>({type:"scatter",mode:"lines+markers",name:s.label,x,y:s.values,line:{color:s.color,width:3},marker:{size:7,color:s.color},hovertemplate:`<b>${s.label}</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>`})),{height:430,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:24,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"GMV",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawCombo(id,barsA,barsB,line){const x=Array.from({length:line.length},(_,i)=>i+1);renderPlot(id,[{type:"bar",name:"站外花费",x,y:barsA,marker:{color:"#93c5fd"},hovertemplate:"<b>站外花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"bar",name:"站内花费",x,y:barsB,marker:{color:"#dbeafe"},hovertemplate:"<b>站内花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"scatter",mode:"lines+markers",name:"店铺GMV",x,y:line,yaxis:"y2",line:{color:"#274472",width:3},marker:{size:7,color:"#274472"},hovertemplate:"<b>店铺GMV</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:58,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:"GMV",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const data=rows.slice(0,state.topN).reverse();const isPct=valueField==="cvr";renderPlot(id,[{type:"bar",orientation:"h",showlegend:false,y:data.map(r=>shortLabel(r[labelField],18)),x:data.map(r=>safeNum(r[valueField])),text:data.map(r=>formatter(r[valueField])),textposition:"outside",cliponaxis:false,marker:{color:"#8db9ff"},customdata:data.map(r=>[r[labelField],safeNum(r[deltaField])]),hovertemplate:`<b>%{customdata[0]}</b><br>${metricAxisTitle(valueField)}：${isPct?"%{x:.1%}":"%{x:,}"}<br>环比：%{customdata[1]:+.1%}<extra></extra>`}],{height:Math.max(390,120+data.length*34),paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:180,r:80,t:16,b:36},hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),xaxis:{title:metricAxisTitle(valueField),gridcolor:"#eef3fb",zeroline:false,tickformat:isPct?".0%":",.0f",tickfont:{color:"#64748b"}},yaxis:{showgrid:false,tickfont:{color:"#334155"},automargin:true}},basePlotConfig());}
    function drawPaired(id,rows){const data=rows.slice(0,state.topN);renderPlot(id,[{type:"bar",name:"Shopee",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.shopee)),marker:{color:"#60a5fa"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>Shopee销量：%{y:,}<extra></extra>"},{type:"bar",name:"TT",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.tt)),marker:{color:"#cfe1ff"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>TT销量：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:60,r:24,t:16,b:90},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"销量",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function bindGlobal(){FILTER_FIELDS.forEach(f=>document.getElementById(`${f}-filter`).addEventListener("change",()=>{state.filters[f]=pick(document.getElementById(`${f}-filter`));renderAll();}));document.getElementById("metric-filter").addEventListener("change",e=>{state.metricFocus=e.target.value;renderAll();});document.getElementById("topn-filter").addEventListener("change",e=>{state.topN=Math.min(50,Math.max(1,Number(e.target.value)||8));renderAll();});document.getElementById("reset-filter").onclick=()=>{FILTER_FIELDS.forEach(f=>state.filters[f]=[]);state.metricFocus="gmv";state.topN=8;renderAll(true);};document.querySelectorAll("[data-download]").forEach(b=>b.onclick=()=>downloadCanvas(b.dataset.download));}
    function renderAll(reset=false){updateFilterControls(reset);const metrics=overallMetrics();renderSummary(metrics);const daily=dailySeries();const focus=getMetricConfig();drawLine("gmv-trend",[{label:"4月当前",color:"#2f6bff",values:daily.gmvCurrent},{label:"3月同期",color:"#84b5ff",values:daily.gmvPrior}]);drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");drawBars("group-chart",groupRows,"label",focus.field,focus.delta,focus.formatter);drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);drawBars("category-chart",categoryRows,"label",focus.field,focus.delta,focus.formatter);drawBars("product-chart",productRows,"label",focus.field,focus.delta,focus.formatter);drawPaired("tt-gap-chart",ttRows);drawBars("bundle-chart",bundleRows,"label",focus.field,focus.delta,focus.formatter);renderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);renderTable("category",categoryDiagnosis);renderTable("product",productDiagnosis);}
    setupMeta();updateFilterControls(true);bindGlobal();renderAll(true);
  </script>
</body>
</html>
"""


def render_html(
    payload: dict[str, Any],
    plotly_src: str = f"./{PLOTLY_JS_NAME}",
    plotly_inline_text: str | None = None,
) -> str:
    html = (
        HTML_TEMPLATE.replace("__TITLE__", payload["meta"]["title"])
        .replace("__REPORT_DATA__", json.dumps(payload, ensure_ascii=False))
        .replace("__PLOTLY_SRC__", plotly_src)
    )
    html = re.sub(
        r'<canvas id="([^"]+)"[^>]*></canvas>',
        r'<div id="\1" class="plot"></div>',
        html,
    )
    html = re.sub(
        r'<section class="filter-panel">[\s\S]*?</section>\s*<section class="summary-grid" id="summary-grid"></section>',
        """
    <section class="filter-panel">
      <div class="filter-head">
        <div>
          <div class="filter-title">筛选控件</div>
          <div class="filter-summary-line" id="filter-summary"></div>
        </div>
        <div class="filter-actions">
          <button class="btn-secondary" id="reset-filter">重置筛选</button>
        </div>
      </div>
      <div class="filter-grid">
        <div class="filter-box"><label for="group-filter">模块视角</label><select id="group-filter"></select></div>
        <div class="filter-box"><label for="metric-filter">趋势关注指标</label><select id="metric-filter"></select></div>
        <div class="filter-box"><label for="category-filter">品类</label><select id="category-filter"></select></div>
        <div class="filter-box"><label for="product-filter">产品</label><select id="product-filter"></select></div>
        <div class="filter-box"><label for="shop-filter">店铺</label><select id="shop-filter"></select></div>
        <div class="filter-box"><label for="topn-filter">Top N</label><select id="topn-filter"></select></div>
      </div>
    </section>
    <section class="summary-grid" id="summary-grid"></section>""",
        html,
        count=1,
    )
    html = html.replace('<select id="group-filter"></select>', '<select id="group-filter" multiple></select>')
    html = html.replace('<label for="category-filter">鍝佺被</label><select id="category-filter"></select>', '<label for="category-filter">鍝佺被</label><input id="category-search" class="filter-search" type="text" placeholder="鎼滅储鍝佺被"><select id="category-filter" multiple></select>')
    html = html.replace('<label for="product-filter">浜у搧</label><select id="product-filter"></select>', '<label for="product-filter">浜у搧</label><input id="product-search" class="filter-search" type="text" placeholder="鎼滅储浜у搧"><select id="product-filter" multiple></select>')
    html = html.replace('<label for="category-filter">品类</label><select id="category-filter"></select>', '<label for="category-filter">品类</label><input id="category-search" class="filter-search" type="text" placeholder="搜索品类"><select id="category-filter" multiple></select>')
    html = html.replace('<label for="product-filter">产品</label><select id="product-filter"></select>', '<label for="product-filter">产品</label><input id="product-search" class="filter-search" type="text" placeholder="搜索产品"><select id="product-filter" multiple></select>')
    html = html.replace('<select id="shop-filter"></select>', '<select id="shop-filter" multiple></select>')
    html = html.replace('<label for="category-filter">鍝佺被</label><input id="category-search" class="filter-search" type="text" placeholder="鎼滅储鍝佺被"><select id="category-filter" multiple></select>', '<label for="category-filter">鍝佺被</label><select id="category-filter" multiple></select>')
    html = html.replace('<label for="product-filter">浜у搧</label><input id="product-search" class="filter-search" type="text" placeholder="鎼滅储浜у搧"><select id="product-filter" multiple></select>', '<label for="product-filter">浜у搧</label><select id="product-filter" multiple></select>')
    html = html.replace('<label for="category-filter">品类</label><input id="category-search" class="filter-search" type="text" placeholder="搜索品类"><select id="category-filter" multiple></select>', '<label for="category-filter">品类</label><select id="category-filter" multiple></select>')
    html = html.replace('<label for="product-filter">产品</label><input id="product-search" class="filter-search" type="text" placeholder="搜索产品"><select id="product-filter" multiple></select>', '<label for="product-filter">产品</label><select id="product-filter" multiple></select>')
    html = html.replace(
        "</style>",
        """
    .filter-search{display:none!important;width:100%;height:40px;border-radius:14px;border:1px solid rgba(47,107,255,.16);padding:0 14px;margin:0 0 8px;background:#fff;color:var(--text);font:inherit}
    .filter-box select[multiple]{height:188px;padding:10px 14px}
    .filter-box select[multiple] option{padding:6px 4px}
    .filter-box select[multiple].native-hidden{display:none}
    .multi-select{position:relative}
    .multi-trigger{width:100%;height:64px;border-radius:20px;border:1px solid rgba(47,107,255,.16);padding:0 20px;background:#fff;color:var(--text);box-shadow:inset 0 1px 0 rgba(255,255,255,.9);display:flex;align-items:center;justify-content:space-between;font:inherit;cursor:pointer}
    .multi-trigger::after{content:"";width:10px;height:10px;border-right:2px solid #64748b;border-bottom:2px solid #64748b;transform:rotate(45deg) translateY(-2px);flex:0 0 auto}
    .multi-trigger-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .multi-panel{position:absolute;top:72px;left:0;right:0;z-index:40;background:#fff;border:1px solid rgba(47,107,255,.16);border-radius:20px;box-shadow:0 18px 40px rgba(15,23,42,.10);padding:12px;display:none}
    .multi-select.open .multi-panel{display:block}
    .multi-search{width:100%;height:40px;border-radius:14px;border:1px solid rgba(47,107,255,.16);padding:0 14px;background:#f8fbff;color:var(--text);font:inherit;margin-bottom:10px}
    .multi-actions{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}
    .multi-mini{border:1px solid rgba(47,107,255,.14);background:#fff;border-radius:999px;padding:6px 10px;color:var(--blue-deep);font:inherit;font-size:12px;cursor:pointer}
    .multi-options{max-height:260px;overflow:auto;padding-right:2px}
    .multi-option{display:grid;grid-template-columns:18px 1fr auto;gap:10px;align-items:center;padding:8px 6px;border-radius:12px}
    .multi-option:hover{background:#f7faff}
    .multi-option input{margin:0}
    .multi-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px;color:#1f2937}
</style>""",
        1,
    )
    html = re.sub(
        r"function renderNarratives\(metrics,groupRows,moduleRows,categoryRows,bundleRows,ttRows\)\{[\s\S]*?function bindGlobal\(\)",
        "function bindGlobal()",
        html,
        count=1,
    )
    html = re.sub(
        r"function renderAll\(reset=false\)\{[\s\S]*?renderTable\(\"product\",productDiagnosis\);\}",
        "",
        html,
        count=1,
    )
    html = re.sub(
        r'const TABLE_COLUMNS = \{[\s\S]*?const state = \{filters:\{shop:\[\],group:\[\],category:\[\],product:\[\]\},metricFocus:"gmv",topN:\d+,tableState:\{category:initTableState\("category"\),product:initTableState\("product"\)\}\};',
        """const TABLE_COLUMNS = {
      category:[["label","品类"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["outerSpend","站外花费"],["impressions","站外曝光量"],["diagnosis","诊断"],["todo","下一步"]],
      product:[["label","产品"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["outerSpend","站外花费"],["impressions","站外曝光量"],["diagnosis","诊断"],["todo","下一步"]]
    };
    const state = {filters:{shop:[],group:[],category:[],product:[]},metricFocus:"gmv",topN:10,tableState:{category:initTableState("category"),product:initTableState("product")}};""",
        html,
        count=1,
    )
    html = re.sub(
        r"function renderTable\(kind,rows\)\{[\s\S]*?bindTableActions\(kind,rows\);\}",
        """function renderTable(kind,rows){const controls=document.getElementById(`${kind}-table-controls`);const holder=document.getElementById(`${kind}-table`);const ts=state.tableState[kind];const columns=TABLE_COLUMNS[kind];controls.innerHTML=`<div class="table-config">${columns.map(([k,l])=>`<div><label>${l}</label><div class="toggle-chip"><input type="checkbox" data-kind="${kind}" data-column="${k}" ${ts.visible[k]?"checked":""}><input type="text" data-kind="${kind}" data-label="${k}" value="${ts.labels[k]}"></div></div>`).join("")}<div><label>行控制</label><button class="mini-btn" data-reset="${kind}">恢复全部隐藏行</button></div><div><label>导出</label><button class="mini-btn" data-export="${kind}">下载当前表格 CSV</button></div></div>`;const visible=columns.filter(([k])=>ts.visible[k]);const list=rows.filter(r=>!ts.hiddenRows.has(r.label)).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,state.topN);if(!list.length){holder.innerHTML=`<div class="empty-state">当前筛选下暂无${kind==="category"?"品类":"产品"}诊断表数据。</div>`;bindTableActions(kind,rows);return;}holder.innerHTML=`<div class="table-wrap"><table><thead><tr>${visible.map(([k])=>`<th>${ts.labels[k]}</th>`).join("")}<th>操作</th></tr></thead><tbody>${list.map(r=>`<tr>${visible.map(([k])=>`<td>${formatCell(k,r[k])}</td>`).join("")}<td><button class="mini-btn" data-hide="${kind}" data-row="${r.label}">隐藏</button></td></tr>`).join("")}</tbody></table></div>`;bindTableActions(kind,rows);}""",
        html,
        count=1,
    )
    html = re.sub(
        r"function exportCsv\(kind,rows\)\{[\s\S]*?URL\.revokeObjectURL\(url\);\}",
        lambda _: """function exportCsv(kind,rows){const cols=TABLE_COLUMNS[kind].filter(([k])=>state.tableState[kind].visible[k]);const header=cols.map(([k])=>state.tableState[kind].labels[k]).join(",");const lines=[header];rows.filter(r=>!state.tableState[kind].hiddenRows.has(r.label)).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,state.topN).forEach(row=>{lines.push(cols.map(([k])=>`"${String(formatCell(k,row[k])).replaceAll('"','""')}"`).join(","));});const blob=new Blob(["\\uFEFF"+lines.join("\\n")],{type:"text/csv;charset=utf-8;"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=`${kind}_diagnosis.csv`;a.click();URL.revokeObjectURL(url);}""",
        html,
        count=1,
    )
    html = html.replace('state.topN=8;renderAll(true);', 'state.topN=10;renderAll(true);')
    html = html.replace(
        "    setupMeta();updateFilterControls(true);bindGlobal();renderAll(true);",
        """
    function shortLabel(text,limit=16){const t=String(text||"");return t.length>limit?t.slice(0,limit)+"…":t;}
    function metricAxisTitle(field){return field==="cvr"?"转化率":field==="visitors"?"访客":field==="orders"?"订单量":field==="innerSpend"?"站内花费":field==="outerSpend"?"站外花费":field==="spend"?"花费":"GMV";}
    function basePlotConfig(){return {responsive:true,displaylogo:false,displayModeBar:false};}
    function basePlotFont(){return {family:"Microsoft YaHei, PingFang SC, Segoe UI, sans-serif",size:12,color:"#64748b"};}
    function renderPlot(id,data,layout,config){const el=document.getElementById(id);try{return Plotly.newPlot(el,data,layout,config);}catch(err){console.error(err);if(el){el.innerHTML=`<div class="empty-state">图表加载失败：${String(err&&err.message?err.message:err)}</div>`;}return null;}}
    function drawLine(id,series){const x=Array.from({length:series[0]?.values?.length||0},(_,i)=>i+1);renderPlot(id,series.map(s=>({type:"scatter",mode:"lines+markers",name:s.label,x,y:s.values,line:{color:s.color,width:3},marker:{size:7,color:s.color},hovertemplate:`<b>${s.label}</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>`})),{height:430,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:24,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"GMV",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawCombo(id,barsA,barsB,line){const x=Array.from({length:line.length},(_,i)=>i+1);renderPlot(id,[{type:"bar",name:"站外花费",x,y:barsA,marker:{color:"#93c5fd"},hovertemplate:"<b>站外花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"bar",name:"站内花费",x,y:barsB,marker:{color:"#dbeafe"},hovertemplate:"<b>站内花费</b><br>日期：%{x}<br>金额：%{y:,}<extra></extra>"},{type:"scatter",mode:"lines+markers",name:"店铺GMV",x,y:line,yaxis:"y2",line:{color:"#274472",width:3},marker:{size:7,color:"#274472"},hovertemplate:"<b>店铺GMV</b><br>日期：%{x}<br>GMV：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:58,t:16,b:86},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"linear",dtick:1,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:"花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:"GMV",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const data=rows.slice(0,state.topN).reverse();const isPct=valueField==="cvr";renderPlot(id,[{type:"bar",orientation:"h",showlegend:false,y:data.map(r=>shortLabel(r[labelField],18)),x:data.map(r=>safeNum(r[valueField])),text:data.map(r=>formatter(r[valueField])),textposition:"outside",cliponaxis:false,marker:{color:"#8db9ff"},customdata:data.map(r=>[r[labelField],safeNum(r[deltaField])]),hovertemplate:`<b>%{customdata[0]}</b><br>${metricAxisTitle(valueField)}：${isPct?"%{x:.1%}":"%{x:,}"}<br>环比：%{customdata[1]:+.1%}<extra></extra>`}],{height:Math.max(390,120+data.length*34),paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:180,r:80,t:16,b:36},hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),xaxis:{title:metricAxisTitle(valueField),gridcolor:"#eef3fb",zeroline:false,tickformat:isPct?".0%":",.0f",tickfont:{color:"#64748b"}},yaxis:{showgrid:false,tickfont:{color:"#334155"},automargin:true}},basePlotConfig());}
    function drawPaired(id,rows){const data=rows.slice(0,state.topN);renderPlot(id,[{type:"bar",name:"Shopee",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.shopee)),marker:{color:"#60a5fa"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>Shopee销量：%{y:,}<extra></extra>"},{type:"bar",name:"TT",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.tt)),marker:{color:"#cfe1ff"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>TT销量：%{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:60,r:24,t:16,b:90},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",y:-0.2,x:0,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"销量",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    setupMeta();updateFilterControls(true);bindGlobal();renderAll(true);""",
    )
    html = html.replace(
        "</body>",
        """
  <script>
    function chartDayLabels(start,count){const base=new Date(`${start}T00:00:00`);return Array.from({length:count},(_,idx)=>{const d=new Date(base);d.setDate(base.getDate()+idx);return `${d.getMonth()+1}/${d.getDate()}`;});}
    function chartLegendLayout(){return {orientation:"h",y:1.16,yanchor:"bottom",x:1,xanchor:"right",bgcolor:"rgba(255,255,255,.96)",bordercolor:"#dbe3ef",borderwidth:1,font:{family:basePlotFont().family,size:12,color:"#334155"}};}
    function chartAxisTitle(text){return {text,standoff:16,font:{family:basePlotFont().family,size:12,color:"#64748b"}};}
    function chartMetricLabel(field){return field==="cvr"?"CVR":field==="visitors"?"Visitors":field==="orders"?"Orders":field==="innerSpend"?"Inner Spend":field==="outerSpend"?"Outer Spend":field==="spend"?"Spend":"GMV";}
    function drawLine(id,series){const count=series[0]?.values?.length||0;const axisLabels=chartDayLabels(DATA.meta.current_start,count);const compareLabels=chartDayLabels(DATA.meta.prior_start,count);renderPlot(id,series.map((s,idx)=>({type:"scatter",mode:"lines+markers",name:s.label,x:axisLabels,y:s.values,customdata:idx===0?axisLabels:compareLabels,line:{color:s.color,width:3},marker:{size:7,color:s.color},hovertemplate:`<b>${s.label}</b><br>Date: %{customdata}<br>GMV: %{y:,}<extra></extra>`})),{height:430,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:92,r:34,t:78,b:102},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:chartLegendLayout(),xaxis:{title:chartAxisTitle("Date"),type:"category",categoryorder:"array",categoryarray:axisLabels,tickangle:-40,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:chartAxisTitle("GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawCombo(id,barsA,barsB,line){const axisLabels=chartDayLabels(DATA.meta.current_start,line.length);renderPlot(id,[{type:"bar",name:"Outer Spend",x:axisLabels,y:barsA,marker:{color:"#93c5fd"},hovertemplate:"<b>Outer Spend</b><br>Date: %{x}<br>Amount: %{y:,}<extra></extra>"},{type:"bar",name:"Inner Spend",x:axisLabels,y:barsB,marker:{color:"#dbeafe"},hovertemplate:"<b>Inner Spend</b><br>Date: %{x}<br>Amount: %{y:,}<extra></extra>"},{type:"scatter",mode:"lines+markers",name:"Shop GMV",x:axisLabels,y:line,yaxis:"y2",line:{color:"#274472",width:3},marker:{size:7,color:"#274472"},hovertemplate:"<b>Shop GMV</b><br>Date: %{x}<br>GMV: %{y:,}<extra></extra>"}],{height:430,barmode:"group",bargap:.18,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:92,r:110,t:78,b:102},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:chartLegendLayout(),xaxis:{title:chartAxisTitle("Date"),type:"category",categoryorder:"array",categoryarray:axisLabels,tickangle:-40,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},yaxis:{title:chartAxisTitle("Spend"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:chartAxisTitle("GMV"),overlaying:"y",side:"right",automargin:true,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());}
    function drawBars(id,rows,labelField,valueField,deltaField,formatter){const data=rows.slice(0,state.topN);const isPct=valueField==="cvr";renderPlot(id,[{type:"bar",showlegend:false,x:data.map(r=>shortLabel(r[labelField],12)),y:data.map(r=>safeNum(r[valueField])),text:data.map(r=>formatter(r[valueField])),textposition:"outside",cliponaxis:false,marker:{color:"#8db9ff"},customdata:data.map(r=>[r[labelField],safeNum(r[deltaField])]),hovertemplate:`<b>%{customdata[0]}</b><br>${chartMetricLabel(valueField)}: ${isPct?"%{y:.1%}":"%{y:,}"}<br>MoM: %{customdata[1]:+.1%}<extra></extra>`}],{height:Math.max(430,360+Math.ceil(data.length/12)*24),paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:84,r:28,t:28,b:132},hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),xaxis:{tickangle:-28,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:chartAxisTitle(chartMetricLabel(valueField)),gridcolor:"#eef3fb",zeroline:false,tickformat:isPct?".0%":",.0f",tickfont:{color:"#64748b"},automargin:true}},basePlotConfig());}
    function drawPaired(id,rows){const data=rows.slice(0,state.topN);renderPlot(id,[{type:"bar",name:"Shopee",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.shopee)),marker:{color:"#60a5fa"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>Shopee Sales: %{y:,}<extra></extra>"},{type:"bar",name:"TT",x:data.map(r=>shortLabel(r.label,10)),y:data.map(r=>safeNum(r.tt)),marker:{color:"#cfe1ff"},customdata:data.map(r=>r.label),hovertemplate:"<b>%{customdata}</b><br>TT Sales: %{y:,}<extra></extra>"}],{height:430,barmode:"group",paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:68,r:24,t:72,b:104},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:chartLegendLayout(),xaxis:{tickangle:-18,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:chartAxisTitle("Sales"),gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"},automargin:true}},basePlotConfig());}
    TABLE_COLUMNS.category=[["label","品类"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["innerSpendMom","环比"],["outerSpend","站外花费"],["outerSpendMom","环比"],["impressions","站外曝光量"],["impressionsMom","环比"],["diagnosis","诊断"],["todo","下一步"]];
    TABLE_COLUMNS.product=[["label","产品"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["innerSpendMom","环比"],["outerSpend","站外花费"],["outerSpendMom","环比"],["impressions","站外曝光量"],["impressionsMom","环比"],["diagnosis","诊断"],["todo","下一步"]];
    state.tableState.category=initTableState("category");
    state.tableState.product=initTableState("product");
    state.topN=state.topN||10;
    function formatCell(key,value){if(["gmv","innerSpend","outerSpend"].includes(key))return fmtMoney(value);if(["visitors","impressions"].includes(key))return fmtNum(value);if(["cvr","gmvMom","visitorsMom","cvrMom","innerSpendMom","outerSpendMom","impressionsMom"].includes(key))return fmtPct(value);return value;}
    function fillSelect(select,options,selected,keyword=""){const selectedSet=new Set(selected||[]);const lower=String(keyword||"").trim().toLowerCase();const filtered=(options||[]).filter(v=>!lower||String(v).toLowerCase().includes(lower));select.innerHTML="";filtered.forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;o.selected=selectedSet.has(v);select.appendChild(o);});if(select.multiple){select.size=Math.min(Math.max(filtered.length,6),10);}}
    function filterSummaryText(values){if(!values||!values.length)return "全部";return values.length<=2?values.join(" / "):`${values.slice(0,2).join(" / ")} 等${values.length}项`;}
    function buildDimensionRows(dim){const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);const rows=Array.from(keys).map(key=>{const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders),pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders),cis=safeNum(current.inner_spend),pis=safeNum(prior.inner_spend),cos=safeNum(current.outer_spend),pos=safeNum(prior.outer_spend),imp=safeNum(current.impressions),pimp=safeNum(prior.impressions);return {label:key,gmv:g,visitors:v,orders:o,cvr:ratio(o,v),innerSpend:cis,outerSpend:cos,impressions:imp,gmvMom:pctChange(g,pg),ordersMom:pctChange(o,poo),visitorsMom:pctChange(v,pv),cvrMom:pctChange(ratio(o,v),ratio(poo,pv)),innerSpendMom:pctChange(cis,pis),outerSpendMom:pctChange(cos,pos),impressionsMom:pctChange(imp,pimp)};}).filter(r=>r.label&&r.label!=="未归类"&&r.label!=="鏈綊绫?");const cfg=getMetricConfig();return rows.sort((a,b)=>safeNum(b[cfg.field])-safeNum(a[cfg.field]));}
    function updateFilterControls(reset=false){fillSelect(document.getElementById("group-filter"),DATA.filters.group,state.filters.group);fillSimpleSelect(document.getElementById("metric-filter"),METRIC_OPTIONS,state.metricFocus);fillSelect(document.getElementById("category-filter"),DATA.filters.category,state.filters.category,document.getElementById("category-search")?.value||"");fillSelect(document.getElementById("product-filter"),DATA.filters.product,state.filters.product,document.getElementById("product-search")?.value||"");fillSelect(document.getElementById("shop-filter"),DATA.filters.shop,state.filters.shop);fillSimpleSelect(document.getElementById("topn-filter"),TOP_N_OPTIONS.map(v=>({value:v,label:String(v)})),state.topN);document.getElementById("filter-summary").textContent=[`模块视角：${filterSummaryText(state.filters.group)}`,`品类：${filterSummaryText(state.filters.category)}`,`产品：${filterSummaryText(state.filters.product)}`,`店铺：${filterSummaryText(state.filters.shop)}`,`Top N：${state.topN}`].join(" ｜ ");}
    function bindGlobal(){FILTER_FIELDS.forEach(f=>document.getElementById(`${f}-filter`).addEventListener("change",()=>{state.filters[f]=pick(document.getElementById(`${f}-filter`));renderAll();}));document.getElementById("metric-filter").addEventListener("change",e=>{state.metricFocus=e.target.value;renderAll();});document.getElementById("topn-filter").addEventListener("change",e=>{state.topN=Math.min(50,Math.max(1,Number(e.target.value)||10));renderAll();});const categorySearch=document.getElementById("category-search");if(categorySearch)categorySearch.addEventListener("input",()=>updateFilterControls(false));const productSearch=document.getElementById("product-search");if(productSearch)productSearch.addEventListener("input",()=>updateFilterControls(false));document.getElementById("reset-filter").onclick=()=>{FILTER_FIELDS.forEach(f=>state.filters[f]=[]);if(categorySearch)categorySearch.value="";if(productSearch)productSearch.value="";state.metricFocus="gmv";state.topN=10;renderAll(true);};document.querySelectorAll("[data-download]").forEach(b=>b.onclick=()=>downloadCanvas(b.dataset.download));}
    if(typeof bindGlobal==="function"){bindGlobal();}
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    TABLE_COLUMNS.category=[["label","品类"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["innerSpendMom","环比"],["outerSpend","站外花费"],["outerSpendMom","环比"],["impressions","站外曝光量"],["impressionsMom","环比"],["diagnosis","诊断"],["todo","下一步"]];
    TABLE_COLUMNS.product=[["label","产品"],["gmv","GMV"],["gmvMom","环比"],["visitors","访客"],["visitorsMom","环比"],["cvr","转化率"],["cvrMom","环比"],["innerSpend","站内花费"],["innerSpendMom","环比"],["outerSpend","站外花费"],["outerSpendMom","环比"],["impressions","站外曝光量"],["impressionsMom","环比"],["diagnosis","诊断"],["todo","下一步"]];
    state.tableState.category=initTableState("category");
    state.tableState.product=initTableState("product");
    state.topN=state.topN||10;
    function formatCell(key,value){if(["gmv","innerSpend","outerSpend"].includes(key))return fmtMoney(value);if(["visitors","impressions"].includes(key))return fmtNum(value);if(["cvr","gmvMom","visitorsMom","cvrMom","innerSpendMom","outerSpendMom","impressionsMom"].includes(key))return fmtPct(value);return value;}
    function renderSummary(metrics){const currentOuterShare=ratio(metrics.current.outer_spend,metrics.current.totalSpend);const priorOuterShare=ratio(metrics.prior.outer_spend,metrics.prior.totalSpend);const currentInnerShare=ratio(metrics.current.inner_spend,metrics.current.totalSpend);const priorInnerShare=ratio(metrics.prior.inner_spend,metrics.prior.totalSpend);const cards=[["店铺GMV",fmtMoney(metrics.current.gmv),pctChange(metrics.current.gmv,metrics.prior.gmv),`对比 ${DATA.meta.prior_start} ~ ${DATA.meta.prior_end}`],["订单量",fmtNum(metrics.current.orders),pctChange(metrics.current.orders,metrics.prior.orders),`AOV ${fmtMoney(metrics.current.aov)} | 环比 ${fmtDelta(pctChange(metrics.current.aov,metrics.prior.aov))}`],["访客",fmtNum(metrics.current.visitors),pctChange(metrics.current.visitors,metrics.prior.visitors),`CTR ${fmtPct(metrics.current.ctr)} | 环比 ${fmtDelta(pctChange(metrics.current.ctr,metrics.prior.ctr))}`],["转化率",fmtPct(metrics.current.cvr),pctChange(metrics.current.cvr,metrics.prior.cvr),`ROI ${fmtNum(metrics.current.roi,2)} | 环比 ${fmtDelta(pctChange(metrics.current.roi,metrics.prior.roi))}`],["站内花费",fmtMoney(metrics.current.inner_spend),pctChange(metrics.current.inner_spend,metrics.prior.inner_spend),`占总花费 ${fmtPct(currentInnerShare)} | 环比 ${fmtDelta(pctChange(currentInnerShare,priorInnerShare))}`],["站外花费",fmtMoney(metrics.current.outer_spend),pctChange(metrics.current.outer_spend,metrics.prior.outer_spend),`站外曝光 ${fmtNum(metrics.current.impressions)} | 环比 ${fmtDelta(pctChange(metrics.current.impressions,metrics.prior.impressions))}`]];document.getElementById("summary-grid").innerHTML=cards.map(([title,value,delta,note])=>`<div class="card summary-card"><div class="section-label">${title}</div><div class="kpi-value">${value}</div><span class="delta ${deltaClass(delta)}">${fmtDelta(delta)}</span><span class="muted">${note}</span></div>`).join("");}
    function filterSummaryText(values){if(!values||!values.length)return "全部";return values.length<=2?values.join(" / "):`${values.slice(0,2).join(" / ")} 等${values.length}项`;}
    function fillSelect(select,options,selected){const selectedSet=new Set(selected||[]);select.innerHTML="";(options||[]).forEach(v=>{const o=document.createElement("option");o.value=v;o.textContent=v;o.selected=selectedSet.has(v);select.appendChild(o);});if(select.multiple)select.classList.add("native-hidden");}
    function buildDimensionRows(dim){const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);const rows=Array.from(keys).map(key=>{const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders),pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders),cis=safeNum(current.inner_spend),pis=safeNum(prior.inner_spend),cos=safeNum(current.outer_spend),pos=safeNum(prior.outer_spend),imp=safeNum(current.impressions),pimp=safeNum(prior.impressions);return {label:key,gmv:g,visitors:v,orders:o,cvr:ratio(o,v),innerSpend:cis,outerSpend:cos,impressions:imp,gmvMom:pctChange(g,pg),ordersMom:pctChange(o,poo),visitorsMom:pctChange(v,pv),cvrMom:pctChange(ratio(o,v),ratio(poo,pv)),innerSpendMom:pctChange(cis,pis),outerSpendMom:pctChange(cos,pos),impressionsMom:pctChange(imp,pimp)};}).filter(r=>r.label&&r.label!=="未归类"&&r.label!=="鏈綊绫?"&&r.label!=="閺堫亜缍婄猾?");const cfg=getMetricConfig();return rows.sort((a,b)=>safeNum(b[cfg.field])-safeNum(a[cfg.field]));}
    const MULTI_CONFIG={group:{search:false},category:{search:true},product:{search:true},shop:{search:false}};
    function syncNativeSelect(select,values){const set=new Set(values||[]);Array.from(select.options).forEach(o=>{o.selected=set.has(o.value);});}
    function ensureMultiSelect(key){const select=document.getElementById(`${key}-filter`);if(!select||!select.multiple)return null;select.classList.add("native-hidden");let wrap=select.previousElementSibling;if(!wrap||!wrap.classList.contains("multi-select")){wrap=document.createElement("div");wrap.className="multi-select";wrap.innerHTML=`<button type="button" class="multi-trigger"><span class="multi-trigger-text">全部</span></button><div class="multi-panel">${MULTI_CONFIG[key].search?'<input class="multi-search" type="text" placeholder="搜索">':''}<div class="multi-actions"><button type="button" class="multi-mini" data-action="clear">清空</button><button type="button" class="multi-mini" data-action="close">收起</button></div><div class="multi-options"></div></div>`;select.parentNode.insertBefore(wrap,select);wrap.querySelector(".multi-trigger").onclick=(e)=>{e.stopPropagation();document.querySelectorAll(".multi-select.open").forEach(node=>{if(node!==wrap)node.classList.remove("open");});wrap.classList.toggle("open");};wrap.querySelector(".multi-panel").addEventListener("click",e=>e.stopPropagation());wrap.querySelector('[data-action="clear"]').onclick=()=>{state.filters[key]=[];syncNativeSelect(select,[]);renderAll();wrap.classList.add("open");};wrap.querySelector('[data-action="close"]').onclick=()=>wrap.classList.remove("open");const search=wrap.querySelector(".multi-search");if(search)search.addEventListener("input",()=>renderMultiSelect(key));}return wrap;}
    function renderMultiSelect(key){const select=document.getElementById(`${key}-filter`);const wrap=ensureMultiSelect(key);if(!select||!wrap)return;fillSelect(select,DATA.filters[key],state.filters[key]);const term=(wrap.querySelector(".multi-search")?.value||"").trim().toLowerCase();const options=(DATA.filters[key]||[]).filter(v=>!term||String(v).toLowerCase().includes(term));const selectedSet=new Set(state.filters[key]||[]);wrap.querySelector(".multi-trigger-text").textContent=filterSummaryText(state.filters[key]);const optionsHolder=wrap.querySelector(".multi-options");optionsHolder.innerHTML=options.map(v=>`<label class="multi-option"><input type="checkbox" value="${String(v).replaceAll('"','&quot;')}" ${selectedSet.has(v)?"checked":""}><span class="multi-label">${v}</span><button type="button" class="multi-mini" data-only="${String(v).replaceAll('"','&quot;')}">仅限</button></label>`).join("")||'<div class="empty-state">没有匹配项</div>';optionsHolder.querySelectorAll('input[type="checkbox"]').forEach(input=>{input.onchange=()=>{const next=new Set(state.filters[key]||[]);if(input.checked)next.add(input.value);else next.delete(input.value);state.filters[key]=Array.from(next);syncNativeSelect(select,state.filters[key]);renderAll();wrap.classList.add("open");};});optionsHolder.querySelectorAll("[data-only]").forEach(btn=>{btn.onclick=()=>{state.filters[key]=[btn.dataset.only];syncNativeSelect(select,state.filters[key]);wrap.classList.remove("open");renderAll();};});}
    function updateFilterControls(){["group","category","product","shop"].forEach(renderMultiSelect);fillSimpleSelect(document.getElementById("metric-filter"),METRIC_OPTIONS,state.metricFocus);fillSimpleSelect(document.getElementById("topn-filter"),TOP_N_OPTIONS.map(v=>({value:v,label:String(v)})),state.topN);document.getElementById("filter-summary").textContent=[`模块视角：${filterSummaryText(state.filters.group)}`,`品类：${filterSummaryText(state.filters.category)}`,`产品：${filterSummaryText(state.filters.product)}`,`店铺：${filterSummaryText(state.filters.shop)}`,`Top N：${state.topN}`].join(" ｜ ");}
    function bindGlobal(){if(!window.__dashboardFilterBound){document.addEventListener("click",()=>document.querySelectorAll(".multi-select.open").forEach(node=>node.classList.remove("open")));window.__dashboardFilterBound=true;}document.getElementById("metric-filter").onchange=e=>{state.metricFocus=e.target.value;renderAll();};document.getElementById("topn-filter").onchange=e=>{state.topN=Math.min(50,Math.max(1,Number(e.target.value)||10));renderAll();};document.getElementById("reset-filter").onclick=()=>{FILTER_FIELDS.forEach(f=>state.filters[f]=[]);document.querySelectorAll(".multi-search").forEach(input=>input.value="");state.metricFocus="gmv";state.topN=10;renderAll();};document.querySelectorAll("[data-download]").forEach(b=>b.onclick=()=>downloadCanvas(b.dataset.download));}
    document.querySelectorAll(".filter-search").forEach(node=>node.remove());
    if(typeof setupMeta==="function"){setupMeta();}
    if(typeof bindGlobal==="function"){bindGlobal();}
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    function getSortedFilterOptions(key){
      if(!["shop","category","product"].includes(key)) return [...(DATA.filters[key]||[])];
      const rows=(DATA.shop_daily||[]).filter(row=>
        row.date>=DATA.meta.current_start &&
        row.date<=DATA.meta.current_end &&
        FILTER_FIELDS.every(field=>field===key || matchField(row[field],state.filters[field]))
      );
      const scoreMap=new Map();
      rows.forEach(row=>{
        const name=row[key];
        if(!name) return;
        scoreMap.set(name,(scoreMap.get(name)||0)+safeNum(row.gmv));
      });
      return [...(DATA.filters[key]||[])].sort((a,b)=>{
        const diff=safeNum(scoreMap.get(b))-safeNum(scoreMap.get(a));
        return Math.abs(diff)>0.0001 ? diff : String(a).localeCompare(String(b),"zh-CN");
      });
    }
    function renderMultiSelect(key){
      const select=document.getElementById(`${key}-filter`);
      const wrap=ensureMultiSelect(key);
      if(!select||!wrap) return;
      const orderedOptions=getSortedFilterOptions(key);
      fillSelect(select,orderedOptions,state.filters[key]);
      const term=(wrap.querySelector(".multi-search")?.value||"").trim().toLowerCase();
      const options=orderedOptions.filter(v=>!term||String(v).toLowerCase().includes(term));
      const selectedSet=new Set(state.filters[key]||[]);
      wrap.querySelector(".multi-trigger-text").textContent=filterSummaryText(state.filters[key]);
      const optionsHolder=wrap.querySelector(".multi-options");
      optionsHolder.innerHTML=options.map(v=>`<label class="multi-option"><input type="checkbox" value="${String(v).replaceAll('"','&quot;')}" ${selectedSet.has(v)?"checked":""}><span class="multi-label">${v}</span><button type="button" class="multi-mini" data-only="${String(v).replaceAll('"','&quot;')}">仅限</button></label>`).join("")||'<div class="empty-state">没有匹配项</div>';
      optionsHolder.querySelectorAll('input[type="checkbox"]').forEach(input=>{
        input.onchange=()=>{
          const next=new Set(state.filters[key]||[]);
          if(input.checked) next.add(input.value); else next.delete(input.value);
          state.filters[key]=Array.from(next);
          syncNativeSelect(select,state.filters[key]);
          renderAll();
          wrap.classList.add("open");
        };
      });
      optionsHolder.querySelectorAll("[data-only]").forEach(btn=>{
        btn.onclick=()=>{
          state.filters[key]=[btn.dataset.only];
          syncNativeSelect(select,state.filters[key]);
          wrap.classList.remove("open");
          renderAll();
        };
      });
    }
    function updateFilterControls(){
      ["group","category","product","shop"].forEach(renderMultiSelect);
      fillSimpleSelect(document.getElementById("metric-filter"),METRIC_OPTIONS,state.metricFocus);
      fillSimpleSelect(document.getElementById("topn-filter"),TOP_N_OPTIONS.map(v=>({value:v,label:String(v)})),state.topN);
      document.getElementById("filter-summary").textContent=[
        `模块视角：${filterSummaryText(state.filters.group)}`,
        `品类：${filterSummaryText(state.filters.category)}`,
        `产品：${filterSummaryText(state.filters.product)}`,
        `店铺：${filterSummaryText(state.filters.shop)}`,
        `Top N：${state.topN}`
      ].join(" ｜ ");
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .table-toolbar{position:relative;display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-top:12px;padding:0}
    .table-toolbar .btn-secondary{padding:12px 22px;border-radius:18px;background:#f4f8ff;border:1px solid rgba(47,107,255,.14);color:var(--blue-deep);font-weight:700}
    .table-toolbar .btn-secondary:hover{background:#edf4ff}
    .column-panel{position:absolute;top:64px;left:0;z-index:30;width:min(100%,520px);padding:18px 20px;background:#fff;border:1px solid rgba(47,107,255,.14);border-radius:24px;box-shadow:0 18px 40px rgba(15,23,42,.10)}
    .column-panel[hidden]{display:none}
    .column-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px 18px}
    .column-item{display:flex;align-items:flex-start;gap:10px;font-size:16px;line-height:1.4;color:#334155}
    .column-item input{margin-top:3px;width:18px;height:18px;accent-color:#1783ff}
    @media (max-width:900px){.column-panel{width:100%;left:0;right:0}.column-grid{grid-template-columns:1fr}}
  </style>
  <script>
    function renderTable(kind,rows){
      const controls=document.getElementById(`${kind}-table-controls`);
      const holder=document.getElementById(`${kind}-table`);
      const ts=state.tableState[kind];
      const columns=TABLE_COLUMNS[kind];
      const panelOpen=!!ts.panelOpen;
      controls.innerHTML=`<div class="table-toolbar"><button class="btn-secondary" type="button" data-column-toggle="${kind}">列设置</button><button class="btn-secondary" type="button" data-reset-columns="${kind}">恢复全部列</button><button class="btn-secondary" type="button" data-reset="${kind}">恢复全部行</button><button class="btn-secondary" type="button" data-export="${kind}">下载 CSV</button><div class="column-panel" id="${kind}-column-panel" ${panelOpen?"":"hidden"}><div class="column-grid">${columns.map(([k])=>`<label class="column-item"><input type="checkbox" data-kind="${kind}" data-column="${k}" ${ts.visible[k]?"checked":""}><span>${ts.labels[k]}</span></label>`).join("")}</div></div></div>`;
      const visible=columns.filter(([k])=>ts.visible[k]);
      const list=rows.filter(r=>!ts.hiddenRows.has(r.label)).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,state.topN);
      if(!list.length){
        holder.innerHTML=`<div class="empty-state">当前筛选下暂无${kind==="category"?"品类":"产品"}诊断表数据。</div>`;
        bindTableActions(kind,rows);
        return;
      }
      holder.innerHTML=`<div class="table-wrap"><table><thead><tr>${visible.map(([k])=>`<th>${ts.labels[k]}</th>`).join("")}<th>操作</th></tr></thead><tbody>${list.map(r=>`<tr>${visible.map(([k])=>`<td>${formatCell(k,r[k])}</td>`).join("")}<td><button class="mini-btn" data-hide="${kind}" data-row="${r.label}">隐藏</button></td></tr>`).join("")}</tbody></table></div>`;
      bindTableActions(kind,rows);
    }
    function bindTableActions(kind,rows){
      const ts=state.tableState[kind];
      const toggle=document.querySelector(`button[data-column-toggle="${kind}"]`);
      if(toggle)toggle.onclick=(e)=>{e.stopPropagation();["category","product"].forEach(name=>{state.tableState[name].panelOpen=false;});ts.panelOpen=!ts.panelOpen;renderTable(kind,rows);};
      const resetColumns=document.querySelector(`button[data-reset-columns="${kind}"]`);
      if(resetColumns)resetColumns.onclick=()=>{TABLE_COLUMNS[kind].forEach(([key])=>{ts.visible[key]=true;});ts.panelOpen=false;renderTable(kind,rows);};
      document.querySelectorAll(`input[data-kind="${kind}"][data-column]`).forEach(input=>{
        input.onchange=(e)=>{ts.visible[input.dataset.column]=e.target.checked;ts.panelOpen=true;renderTable(kind,rows);};
      });
      document.querySelectorAll(`button[data-hide="${kind}"]`).forEach(btn=>btn.onclick=()=>{ts.hiddenRows.add(btn.dataset.row);renderTable(kind,rows);});
      const reset=document.querySelector(`button[data-reset="${kind}"]`);
      if(reset)reset.onclick=()=>{ts.hiddenRows.clear();renderTable(kind,rows);};
      const exp=document.querySelector(`button[data-export="${kind}"]`);
      if(exp)exp.onclick=()=>exportCsv(kind,rows);
      const panel=document.getElementById(`${kind}-column-panel`);
      if(panel)panel.onclick=(e)=>e.stopPropagation();
    }
    if(!window.__tablePanelCloserBound){
      document.addEventListener("click",()=>{
        let changed=false;
        ["category","product"].forEach(kind=>{
          if(state.tableState[kind]&&state.tableState[kind].panelOpen){
            state.tableState[kind].panelOpen=false;
            changed=true;
          }
        });
        if(changed){
          const categoryRows=buildDiagnosisRows("category");
          const productRows=buildDiagnosisRows("product");
          renderTable("category",categoryRows);
          renderTable("product",productRows);
        }
      });
      window.__tablePanelCloserBound=true;
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        '<div class="narrative" id="group-narrative"></div></div>\n      <div class="card"><div class="chart-head"><div><div class="section-label">Module</div>',
        '<div class="narrative" id="group-narrative"></div></div>\n      <div class="card"><div class="chart-head"><div><div class="section-label">Group Compare</div><h2>分组当前 vs 上期</h2><div class="muted">柱：店铺GMV；线：CVR 与访客</div></div><div class="chart-tools"><button data-download="group-compare-chart">下载图表</button></div></div><div id="group-compare-chart" class="plot"></div><div class="narrative" id="group-compare-narrative"></div></div>\n      <div class="card"><div class="chart-head"><div><div class="section-label">Module</div>',
        1,
    )
    html = html.replace(
        '<div class="section-label">Group</div><h2>分组分析</h2><div class="muted">看 GMV 拉动和效率变化</div>',
        '<div class="section-label">Group</div><h2>分组花费与承接</h2><div class="muted">柱：站内外花费；线：访客与转化率</div>',
        1,
    )
    html = html.replace(
        '<div class="section-label">Group</div><h2>分组花费与承接</h2><div class="muted">柱：站内外花费；线：访客与转化率</div>',
        '<div class="section-label">Group</div><h2>分组花费与承接</h2><div class="muted">柱：站内外花费；线：访客、转化率与GMV环比</div>',
        1,
    )
    html = html.replace(
        '<div class="section-label">Group Compare</div><h2>分组当前 vs 上期</h2><div class="muted">柱：店铺GMV；线：CVR 与访客</div>',
        '<div class="section-label">Group Compare</div><h2>分组当前 vs 上期</h2><div class="muted">柱：4月与3月同期店铺GMV</div>',
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    const __rawBaseRenderNarratives = typeof renderNarratives === "function" ? renderNarratives : null;
    const __legacyNarrativeStubIds = [
      "overall-narrative","spend-narrative","group-narrative","module-narrative","category-narrative","product-narrative",
      "tt-gap-narrative","highlight-box","weakness-box","bundle-box","bundle-narrative","action-box","data-note-box",
      "category-table-controls","category-table","product-table-controls","product-table","category_roi-table-controls",
      "category_roi-table","product_roi-table-controls","product_roi-table"
    ];
    function ensureLegacyStubNode(id){
      const existing = document.getElementById(id);
      if(existing) return null;
      const stub = document.createElement("div");
      stub.id = id;
      stub.className = "analysis-hidden-legacy";
      stub.style.display = "none";
      (document.body || document.documentElement).appendChild(stub);
      return stub;
    }
    function __safeBaseRenderNarratives(){
      if(!__rawBaseRenderNarratives) return null;
      const stubs = __legacyNarrativeStubIds.map(ensureLegacyStubNode).filter(Boolean);
      try{
        return __rawBaseRenderNarratives.apply(this, arguments);
      }finally{
        stubs.forEach((stub)=>stub.remove());
      }
    }
    renderNarratives = __safeBaseRenderNarratives;
    const __baseRenderNarratives = __safeBaseRenderNarratives;
    function buildDimensionRows(dim){
      const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);
      const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);
      const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);
      const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);
      const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);
      const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);
      const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);
      const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);
      const rows=Array.from(keys).map(key=>{
        const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};
        const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};
        const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders);
        const pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders);
        const cis=safeNum(current.inner_spend),pis=safeNum(prior.inner_spend),cos=safeNum(current.outer_spend),pos=safeNum(prior.outer_spend),imp=safeNum(current.impressions),pimp=safeNum(prior.impressions);
        const currentCvr=ratio(o,v),priorCvr=ratio(poo,pv);
        return {
          label:key,
          gmv:g,
          priorGmv:pg,
          visitors:v,
          priorVisitors:pv,
          orders:o,
          priorOrders:poo,
          cvr:currentCvr,
          priorCvr:priorCvr,
          innerSpend:cis,
          outerSpend:cos,
          impressions:imp,
          gmvMom:pctChange(g,pg),
          ordersMom:pctChange(o,poo),
          visitorsMom:pctChange(v,pv),
          cvrMom:pctChange(currentCvr,priorCvr),
          innerSpendMom:pctChange(cis,pis),
          outerSpendMom:pctChange(cos,pos),
          impressionsMom:pctChange(imp,pimp)
        };
      }).filter(r=>r.label&&r.label!=="未归类"&&r.label!=="鏈綊绫?"&&r.label!=="閺堫亜缍婄猾?"&&!(dim==="group"&&r.label==="其他"));
      const cfg=getMetricConfig();
      return rows.sort((a,b)=>safeNum(b[cfg.field])-safeNum(a[cfg.field]));
    }
    function drawGroupCompare(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      renderPlot(id,[
        {type:"bar",name:"4月店铺GMV",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.gmv)),customdata:data.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>4月店铺GMV：%{y:,}<extra></extra>"},
        {type:"bar",name:"3月同期店铺GMV",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorGmv)),customdata:data.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>3月同期店铺GMV：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月CVR",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.cvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>4月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月CVR",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorCvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#8b5cf6",width:3,dash:"dot"},marker:{size:8,color:"#8b5cf6"},hovertemplate:"<b>%{customdata}</b><br>3月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月访客",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.visitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#0ea5e9",width:3},marker:{size:8,color:"#0ea5e9"},hovertemplate:"<b>%{customdata}</b><br>4月访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月访客",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorVisitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#94a3b8",width:3,dash:"dot"},marker:{size:8,color:"#94a3b8"},hovertemplate:"<b>%{customdata}</b><br>3月访客：%{y:,}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:92,r:230,t:96,b:110},
        hovermode:"x unified",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:chartLegendLayout(),
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.8],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("店铺GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"CVR",standoff:8,font:{family:basePlotFont().family,size:12,color:"#16a34a"}},overlaying:"y",anchor:"free",side:"right",position:.88,showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#16a34a",size:11}},
        yaxis3:{title:chartAxisTitle("访客"),overlaying:"y",side:"right",position:1,automargin:true,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#0ea5e9"}}
      },basePlotConfig());
    }
    function renderGroupCompareNarrative(groupRows){
      const holder=document.getElementById("group-compare-narrative");
      if(!holder) return;
      const rows=[...groupRows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      if(!rows.length){
        holder.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      if(rows.length===1){
        const row=rows[0];
        holder.innerHTML=`<p>${row.label} 当前是筛选后的唯一有效分组，4月店铺GMV ${fmtMoney(row.gmv)}，较3月同期 ${fmtDelta(row.gmvMom)}。</p><p>该分组当前访客 ${fmtNum(row.visitors)}、CVR ${fmtPct(row.cvr)}，更适合作为单分组承接效率复盘入口。</p><p><strong>下一步：</strong></p>${narrativeList([`继续细拆 ${row.label} 内部产品与套组，确认增长来自流量放量还是转化改善。`,`同步复核 ${row.label} 的素材命中率与链接承接，避免单分组放量后效率回落。`])}`;
        return;
      }
      const top=rows[0];
      const growth=[...rows].sort((a,b)=>(safeNum(b.gmv)-safeNum(b.priorGmv))-(safeNum(a.gmv)-safeNum(a.priorGmv)))[0];
      const bestCvr=[...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
      const weakTraffic=[...rows].sort((a,b)=>safeNum(a.visitorsMom)-safeNum(b.visitorsMom))[0];
      holder.innerHTML=`<p>${top.label} 仍是当前最大分组，4月店铺GMV ${fmtMoney(top.gmv)}，对比3月同期 ${fmtDelta(top.gmvMom)}；${growth.label} 是本期GMV增量更明显的分组，说明当前增长主要由该分组拉动。</p><p>${bestCvr.label} 当前CVR最高，为 ${fmtPct(bestCvr.cvr)}；${weakTraffic.label} 的访客环比 ${fmtDelta(weakTraffic.visitorsMom)}，更需要优先排查流量缩量或素材疲劳。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续承担主分组保量，同时把预算和素材优先向高转化链接集中。`,`对 ${weakTraffic.label} 先补流量入口，再看是否需要同步修正价格力和详情页承接。`])}`;
    }
    function renderAll(reset=false){
      updateFilterControls(reset);
      const metrics=overallMetrics();
      renderSummary(metrics);
      const daily=dailySeries();
      const focus=getMetricConfig();
      drawLine("gmv-trend",[{label:"4月当前",color:"#2f6bff",values:daily.gmvCurrent},{label:"3月同期",color:"#84b5ff",values:daily.gmvPrior}]);
      drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);
      const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");
      drawBars("group-chart",groupRows,"label",focus.field,focus.delta,focus.formatter);
      drawGroupCompare("group-compare-chart",groupRows);
      drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);
      drawBars("category-chart",categoryRows,"label",focus.field,focus.delta,focus.formatter);
      drawBars("product-chart",productRows,"label",focus.field,focus.delta,focus.formatter);
      drawPaired("tt-gap-chart",ttRows);
      drawBars("bundle-chart",bundleRows,"label",focus.field,focus.delta,focus.formatter);
      if(__baseRenderNarratives) __baseRenderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);
      renderGroupCompareNarrative(groupRows);
      renderTable("category",categoryDiagnosis);
      renderTable("product",productDiagnosis);
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    function drawGroupCompare(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      renderPlot(id,[
        {type:"bar",name:"4月店铺GMV",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.gmv)),customdata:data.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>4月店铺GMV：%{y:,}<extra></extra>"},
        {type:"bar",name:"3月同期店铺GMV",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorGmv)),customdata:data.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>3月同期店铺GMV：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月CVR",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.cvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>4月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月CVR",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorCvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#8b5cf6",width:3,dash:"dot"},marker:{size:8,color:"#8b5cf6"},hovertemplate:"<b>%{customdata}</b><br>3月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月访客",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.visitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#0ea5e9",width:3},marker:{size:8,color:"#0ea5e9"},hovertemplate:"<b>%{customdata}</b><br>4月访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月访客",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.priorVisitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#94a3b8",width:3,dash:"dot"},marker:{size:8,color:"#94a3b8"},hovertemplate:"<b>%{customdata}</b><br>3月访客：%{y:,}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:92,r:230,t:96,b:110},
        hovermode:"x unified",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:chartLegendLayout(),
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.8],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("店铺GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"CVR",standoff:8,font:{family:basePlotFont().family,size:12,color:"#16a34a"}},overlaying:"y",anchor:"free",side:"right",position:.88,showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#16a34a",size:11}},
        yaxis3:{title:{text:"访客",standoff:16,font:{family:basePlotFont().family,size:12,color:"#0ea5e9"}},overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#0ea5e9",size:11}}
      },basePlotConfig());
    }
    function drawGroupSpendMix(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
        return;
      }
      renderPlot(id,[
        {type:"bar",name:"站外花费",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.outerSpend)),customdata:data.map(r=>r.label),marker:{color:"#93c5fd"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
        {type:"bar",name:"站内花费",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.innerSpend)),customdata:data.map(r=>r.label),marker:{color:"#dbeafe"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"访客",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#0ea5e9",width:3},marker:{size:8,color:"#0ea5e9"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"CVR",x:data.map(r=>shortLabel(r.label,12)),y:data.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:92,r:220,t:96,b:110},
        hovermode:"x unified",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:chartLegendLayout(),
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.82],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"访客",standoff:10,font:{family:basePlotFont().family,size:12,color:"#0ea5e9"}},overlaying:"y",anchor:"free",side:"right",position:.9,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#0ea5e9",size:11}},
        yaxis3:{title:{text:"CVR",standoff:18,font:{family:basePlotFont().family,size:12,color:"#16a34a"}},overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#16a34a",size:11}}
      },basePlotConfig());
    }
    function renderGroupSpendNarrative(groupRows){
      const holder=document.getElementById("group-narrative");
      if(!holder) return;
      const rows=[...groupRows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum((b.outerSpend+b.innerSpend)) - safeNum((a.outerSpend+a.innerSpend)));
      if(!rows.length){
        holder.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
        return;
      }
      const topSpend=rows[0];
      const bestCvr=[...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
      const weakTraffic=[...rows].sort((a,b)=>safeNum(a.visitors)-safeNum(b.visitors))[0];
      holder.innerHTML=`<p>${topSpend.label} 是当前花费最高的分组，站内外合计花费 ${fmtMoney(topSpend.innerSpend+topSpend.outerSpend)}，其中站外 ${fmtMoney(topSpend.outerSpend)}、站内 ${fmtMoney(topSpend.innerSpend)}。</p><p>${bestCvr.label} 当前 CVR 最高，为 ${fmtPct(bestCvr.cvr)}；${weakTraffic.label} 访客规模最小，当前更需要优先补流量入口，再看转化承接。</p><p><strong>下一步：</strong></p>${narrativeList([`${topSpend.label} 继续做花费结构复盘，避免站内外投放结构失衡而拉高成本。`,`对 ${weakTraffic.label} 先补访客入口，再看是否要同步修正价格力和详情页承接。`])}`;
    }
    const __patchedBaseRenderAll = renderAll;
    function renderAll(reset=false){
      updateFilterControls(reset);
      const metrics=overallMetrics();
      renderSummary(metrics);
      const daily=dailySeries();
      const focus=getMetricConfig();
      drawLine("gmv-trend",[{label:"4月当前",color:"#2f6bff",values:daily.gmvCurrent},{label:"3月同期",color:"#84b5ff",values:daily.gmvPrior}]);
      drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);
      const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");
      drawGroupSpendMix("group-chart",groupRows);
      drawGroupCompare("group-compare-chart",groupRows);
      drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);
      drawBars("category-chart",categoryRows,"label",focus.field,focus.delta,focus.formatter);
      drawBars("product-chart",productRows,"label",focus.field,focus.delta,focus.formatter);
      drawPaired("tt-gap-chart",ttRows);
      drawBars("bundle-chart",bundleRows,"label",focus.field,focus.delta,focus.formatter);
      if(__baseRenderNarratives) __baseRenderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);
      renderGroupSpendNarrative(groupRows);
      renderGroupCompareNarrative(groupRows);
      renderTable("category",categoryDiagnosis);
      renderTable("product",productDiagnosis);
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    function drawGroupCompare(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      const labels=data.map(r=>shortLabel(r.label,12));
      renderPlot(id,[
        {type:"bar",name:"4月店铺GMV",x:labels,y:data.map(r=>safeNum(r.gmv)),customdata:data.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>4月店铺GMV：%{y:,}<extra></extra>"},
        {type:"bar",name:"3月同期店铺GMV",x:labels,y:data.map(r=>safeNum(r.priorGmv)),customdata:data.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>3月同期店铺GMV：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月CVR",x:labels,y:data.map(r=>safeNum(r.cvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>4月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月CVR",x:labels,y:data.map(r=>safeNum(r.priorCvr)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#8b5cf6",width:3,dash:"dot"},marker:{size:8,color:"#8b5cf6"},hovertemplate:"<b>%{customdata}</b><br>3月CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"4月访客",x:labels,y:data.map(r=>safeNum(r.visitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#0ea5e9",width:3},marker:{size:8,color:"#0ea5e9"},hovertemplate:"<b>%{customdata}</b><br>4月访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"3月访客",x:labels,y:data.map(r=>safeNum(r.priorVisitors)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#94a3b8",width:3,dash:"dot"},marker:{size:8,color:"#94a3b8"},hovertemplate:"<b>%{customdata}</b><br>3月访客：%{y:,}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:92,r:156,t:120,b:118},
        hovermode:"closest",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.2,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.84],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("店铺GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"CVR",standoff:10,font:{family:basePlotFont().family,size:12,color:"#16a34a"}},overlaying:"y",anchor:"free",side:"right",position:.9,showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#16a34a",size:11}},
        yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:",.0f"}
      },basePlotConfig());
    }
    function drawGroupSpendMix(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
        return;
      }
      const labels=data.map(r=>shortLabel(r.label,12));
      renderPlot(id,[
        {type:"bar",name:"站外花费",x:labels,y:data.map(r=>safeNum(r.outerSpend)),customdata:data.map(r=>r.label),marker:{color:"#93c5fd"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
        {type:"bar",name:"站内花费",x:labels,y:data.map(r=>safeNum(r.innerSpend)),customdata:data.map(r=>r.label),marker:{color:"#dbeafe"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:data.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#0ea5e9",width:3},marker:{size:8,color:"#0ea5e9"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"CVR",x:labels,y:data.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:92,r:156,t:120,b:118},
        hovermode:"closest",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.2,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.84],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"访客",standoff:10,font:{family:basePlotFont().family,size:12,color:"#0ea5e9"}},overlaying:"y",anchor:"free",side:"right",position:.9,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#0ea5e9",size:11}},
        yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:".0%"}
      },basePlotConfig());
    }
    function renderAll(reset=false){
      updateFilterControls(reset);
      const metrics=overallMetrics();
      renderSummary(metrics);
      const daily=dailySeries();
      const focus=getMetricConfig();
      drawLine("gmv-trend",[{label:"4月当前",color:"#2f6bff",values:daily.gmvCurrent},{label:"3月同期",color:"#84b5ff",values:daily.gmvPrior}]);
      drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);
      const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");
      drawGroupSpendMix("group-chart",groupRows);
      drawGroupCompare("group-compare-chart",groupRows);
      drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);
      drawBars("category-chart",categoryRows,"label",focus.field,focus.delta,focus.formatter);
      drawBars("product-chart",productRows,"label",focus.field,focus.delta,focus.formatter);
      drawPaired("tt-gap-chart",ttRows);
      drawBars("bundle-chart",bundleRows,"label",focus.field,focus.delta,focus.formatter);
      if(typeof __baseRenderNarratives==="function") __baseRenderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);
      renderGroupSpendNarrative(groupRows);
      renderGroupCompareNarrative(groupRows);
      renderTable("category",categoryDiagnosis);
      renderTable("product",productDiagnosis);
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    function drawGroupSpendMix(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
        return;
      }
      const labels=data.map(r=>shortLabel(r.label,12));
      renderPlot(id,[
        {type:"bar",name:"站外花费",x:labels,y:data.map(r=>safeNum(r.outerSpend)),customdata:data.map(r=>r.label),marker:{color:"#93c5fd"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
        {type:"bar",name:"站内花费",x:labels,y:data.map(r=>safeNum(r.innerSpend)),customdata:data.map(r=>r.label),marker:{color:"#dbeafe"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:data.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#1d9bf0",width:3},marker:{size:8,color:"#1d9bf0"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"CVR",x:labels,y:data.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"},
        {type:"scatter",mode:"lines+markers",name:"GMV环比",x:labels,y:data.map(r=>safeNum(r.gmvMom)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#7c3aed",width:3,dash:"dot"},marker:{size:8,color:"#7c3aed"},hovertemplate:"<b>%{customdata}</b><br>GMV环比：%{y:.1%}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:88,r:92,t:118,b:112},
        hovermode:"x unified",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
        xaxis:{title:chartAxisTitle("分组"),domain:[0,0.93],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
        yaxis2:{title:{text:"访客",standoff:8,font:{family:basePlotFont().family,size:12,color:"#1d9bf0"}},overlaying:"y",anchor:"free",side:"right",position:.985,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#1d9bf0",size:11}},
        yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:".0%"}
      },basePlotConfig());
    }
    function drawGroupCompare(id,rows){
      const data=[...rows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,Math.max(1,state.topN));
      if(!data.length){
        const el=document.getElementById(id);
        if(el) el.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      const labels=data.map(r=>shortLabel(r.label,12));
      renderPlot(id,[
        {type:"bar",name:"4月店铺GMV",x:labels,y:data.map(r=>safeNum(r.gmv)),customdata:data.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>4月店铺GMV：%{y:,}<extra></extra>"},
        {type:"bar",name:"3月同期店铺GMV",x:labels,y:data.map(r=>safeNum(r.priorGmv)),customdata:data.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>3月同期店铺GMV：%{y:,}<extra></extra>"}
      ],{
        height:460,
        barmode:"group",
        bargap:.18,
        paper_bgcolor:"rgba(0,0,0,0)",
        plot_bgcolor:"#ffffff",
        margin:{l:88,r:28,t:118,b:112},
        hovermode:"x unified",
        hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
        font:basePlotFont(),
        legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
        xaxis:{title:chartAxisTitle("分组"),domain:[0,1],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
        yaxis:{title:chartAxisTitle("店铺GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
      },basePlotConfig());
    }
    function renderGroupSpendNarrative(groupRows){
      const holder=document.getElementById("group-narrative");
      if(!holder) return;
      const rows=[...groupRows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum((b.outerSpend+b.innerSpend)) - safeNum((a.outerSpend+a.innerSpend)));
      if(!rows.length){
        holder.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
        return;
      }
      const topSpend=rows[0];
      const bestCvr=[...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
      const bestGrowth=[...rows].sort((a,b)=>safeNum(b.gmvMom)-safeNum(a.gmvMom))[0];
      holder.innerHTML=`<p>${topSpend.label} 是当前花费最高的分组，站内外合计花费 ${fmtMoney(topSpend.innerSpend+topSpend.outerSpend)}，其中站外 ${fmtMoney(topSpend.outerSpend)}、站内 ${fmtMoney(topSpend.innerSpend)}。</p><p>${bestCvr.label} 当前 CVR 最高，为 ${fmtPct(bestCvr.cvr)}；${bestGrowth.label} 的 GMV 环比 ${fmtDelta(bestGrowth.gmvMom)}，说明该分组本期增长弹性更强。</p><p><strong>下一步：</strong></p>${narrativeList([`${topSpend.label} 继续做花费结构复盘，避免站内外投放结构失衡而拉高成本。`,`优先复盘 ${bestGrowth.label} 的增长来源，判断是流量放大还是承接改善后再决定是否继续加预算。`])}`;
    }
    function renderGroupCompareNarrative(groupRows){
      const holder=document.getElementById("group-compare-narrative");
      if(!holder) return;
      const rows=[...groupRows].filter(r=>r.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      if(!rows.length){
        holder.innerHTML='<div class="empty-state">当前筛选下暂无分组对比数据。</div>';
        return;
      }
      if(rows.length===1){
        const row=rows[0];
        holder.innerHTML=`<p>${row.label} 当前是筛选后的唯一有效分组，4月店铺GMV ${fmtMoney(row.gmv)}，较3月同期 ${fmtDelta(row.gmvMom)}。</p><p><strong>下一步：</strong></p>${narrativeList([`继续细拆 ${row.label} 内部产品和套组，确认 GMV 变化是来自价格带、素材还是流量结构变化。`])}`;
        return;
      }
      const top=rows[0];
      const growth=[...rows].sort((a,b)=>safeNum(b.gmvMom)-safeNum(a.gmvMom))[0];
      const weak=[...rows].sort((a,b)=>safeNum(a.gmvMom)-safeNum(b.gmvMom))[0];
      holder.innerHTML=`<p>${top.label} 仍是当前最大分组，4月店铺GMV ${fmtMoney(top.gmv)}，对比3月同期 ${fmtDelta(top.gmvMom)}；${growth.label} 是本期环比表现最强的分组。</p><p>${weak.label} 的 GMV 环比 ${fmtDelta(weak.gmvMom)}，说明它仍是当前分组对比里更需要优先修复的拖累点。</p><p><strong>下一步：</strong></p>${narrativeList([`${top.label} 继续承担主分组保量，优先稳住主销链接和套组承接。`,`针对 ${weak.label} 先排查低效流量和链接承接，再决定是否继续保留当前预算。`])}`;
    }
    function renderAll(reset=false){
      updateFilterControls(reset);
      const metrics=overallMetrics();
      renderSummary(metrics);
      const daily=dailySeries();
      const focus=getMetricConfig();
      drawLine("gmv-trend",[{label:"4月当前",color:"#2f6bff",values:daily.gmvCurrent},{label:"3月同期",color:"#84b5ff",values:daily.gmvPrior}]);
      drawCombo("spend-mix",daily.outerSpend,daily.innerSpend,daily.gmvCurrent);
      const groupRows=buildDimensionRows("group"),moduleRows=buildModuleRows(),categoryRows=buildDimensionRows("category"),productRows=buildDimensionRows("product"),bundleRows=buildBundleRows(),ttRows=buildTTGapRows(),categoryDiagnosis=buildDiagnosisRows("category"),productDiagnosis=buildDiagnosisRows("product");
      drawGroupSpendMix("group-chart",groupRows);
      drawGroupCompare("group-compare-chart",groupRows);
      drawBars("module-chart",moduleRows,"label","spend","spendMom",fmtMoney);
      drawBars("category-chart",categoryRows,"label",focus.field,focus.delta,focus.formatter);
      drawBars("product-chart",productRows,"label",focus.field,focus.delta,focus.formatter);
      drawPaired("tt-gap-chart",ttRows);
      drawBars("bundle-chart",bundleRows,"label",focus.field,focus.delta,focus.formatter);
      if(typeof __baseRenderNarratives==="function") __baseRenderNarratives(metrics,groupRows,moduleRows,categoryRows,productRows,bundleRows,ttRows);
      renderGroupSpendNarrative(groupRows);
      renderGroupCompareNarrative(groupRows);
      renderTable("category",categoryDiagnosis);
      renderTable("product",productDiagnosis);
    }
    if(typeof renderAll==="function"){renderAll(true);}
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      function latestGroupRows(){
        if(typeof buildDimensionRows !== "function") return [];
        return buildDimensionRows("group").filter(r=>r && r.label !== "其他").slice(0,Math.max(1,state.topN));
      }
      function syncGroupHeads(){
        const leftCard = document.getElementById("group-chart");
        const rightCard = document.getElementById("group-compare-chart");
        if(leftCard){
          const muted = leftCard.closest(".card")?.querySelector(".muted");
          if(muted) muted.textContent = "柱：站内外花费；线：访客、转化率与GMV环比";
        }
        if(rightCard){
          const muted = rightCard.closest(".card")?.querySelector(".muted");
          if(muted) muted.textContent = "柱：4月与3月同期店铺GMV";
        }
      }
      function redrawGroupSpend(){
        const data = latestGroupRows();
        if(!data.length || typeof renderPlot !== "function") return;
        const labels = data.map(r=>shortLabel(r.label,12));
        renderPlot("group-chart",[
          {type:"bar",name:"站外花费",x:labels,y:data.map(r=>safeNum(r.outerSpend)),customdata:data.map(r=>r.label),marker:{color:"#93c5fd"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
          {type:"bar",name:"站内花费",x:labels,y:data.map(r=>safeNum(r.innerSpend)),customdata:data.map(r=>r.label),marker:{color:"#dbeafe"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:data.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#1d9bf0",width:3},marker:{size:8,color:"#1d9bf0"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"CVR",x:labels,y:data.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"GMV环比",x:labels,y:data.map(r=>safeNum(r.gmvMom)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#7c3aed",width:3,dash:"dot"},marker:{size:8,color:"#7c3aed"},hovertemplate:"<b>%{customdata}</b><br>GMV环比：%{y:.1%}<extra></extra>"}
        ],{
          height:460,
          barmode:"group",
          bargap:.18,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:88,r:92,t:118,b:112},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
          xaxis:{title:chartAxisTitle("分组"),domain:[0,0.93],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
          yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:{text:"访客",standoff:8,font:{family:basePlotFont().family,size:12,color:"#1d9bf0"}},overlaying:"y",anchor:"free",side:"right",position:.985,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#1d9bf0",size:11}},
          yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:".0%"}
        },basePlotConfig());
      }
      function redrawGroupCompare(){
        const data = latestGroupRows();
        if(!data.length || typeof renderPlot !== "function") return;
        const labels = data.map(r=>shortLabel(r.label,12));
        renderPlot("group-compare-chart",[
          {type:"bar",name:"4月店铺GMV",x:labels,y:data.map(r=>safeNum(r.gmv)),customdata:data.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>4月店铺GMV：%{y:,}<extra></extra>"},
          {type:"bar",name:"3月同期店铺GMV",x:labels,y:data.map(r=>safeNum(r.priorGmv)),customdata:data.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>3月同期店铺GMV：%{y:,}<extra></extra>"}
        ],{
          height:460,
          barmode:"group",
          bargap:.18,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:88,r:28,t:118,b:112},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
          xaxis:{title:chartAxisTitle("分组"),domain:[0,1],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
          yaxis:{title:chartAxisTitle("店铺GMV"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }
      function redrawGroupCards(){
        syncGroupHeads();
        redrawGroupSpend();
        redrawGroupCompare();
      }
      const prevRenderAll = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAll === "function") prevRenderAll(reset);
        redrawGroupCards();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(redrawGroupCards,0); });
      }else{
        setTimeout(redrawGroupCards,0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      function redrawLeftGroupChartFinal(){
        if(typeof buildDimensionRows !== "function" || typeof renderPlot !== "function") return;
        const data = buildDimensionRows("group").filter(r=>r && r.label !== "其他").slice(0,Math.max(1,state.topN));
        if(!data.length) return;
        const labels = data.map(r=>shortLabel(r.label,12));
        const card = document.getElementById("group-chart")?.closest(".card");
        const muted = card?.querySelector(".muted");
        if(muted) muted.textContent = "柱：站内外花费；线：访客、转化率及其环比";
        renderPlot("group-chart",[
          {type:"bar",name:"站外花费",x:labels,y:data.map(r=>safeNum(r.outerSpend)),customdata:data.map(r=>r.label),marker:{color:"#93c5fd"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
          {type:"bar",name:"站内花费",x:labels,y:data.map(r=>safeNum(r.innerSpend)),customdata:data.map(r=>r.label),marker:{color:"#dbeafe"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:data.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:data.map(r=>r.label),line:{color:"#1d9bf0",width:3},marker:{size:8,color:"#1d9bf0"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"CVR",x:labels,y:data.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"访客环比",x:labels,y:data.map(r=>safeNum(r.visitorsMom)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#7c3aed",width:3,dash:"dot"},marker:{size:8,color:"#7c3aed"},hovertemplate:"<b>%{customdata}</b><br>访客环比：%{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"CVR环比",x:labels,y:data.map(r=>safeNum(r.cvrMom)),yaxis:"y3",customdata:data.map(r=>r.label),line:{color:"#ef4444",width:3,dash:"dash"},marker:{size:8,color:"#ef4444"},hovertemplate:"<b>%{customdata}</b><br>CVR环比：%{y:.1%}<extra></extra>"}
        ],{
          height:460,
          barmode:"group",
          bargap:.18,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:88,r:92,t:118,b:112},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
          xaxis:{title:chartAxisTitle("分组"),domain:[0,0.93],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
          yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:{text:"访客",standoff:8,font:{family:basePlotFont().family,size:12,color:"#1d9bf0"}},overlaying:"y",anchor:"free",side:"right",position:.985,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#1d9bf0",size:11}},
          yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:".0%"}
        },basePlotConfig());
      }
      function renderLeftGroupNarrativeFinal(){
        if(typeof buildDimensionRows !== "function") return;
        const holder = document.getElementById("group-narrative");
        if(!holder) return;
        const rows = buildDimensionRows("group").filter(r=>r && r.label !== "其他").sort((a,b)=>safeNum((b.outerSpend+b.innerSpend)) - safeNum((a.outerSpend+a.innerSpend)));
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
          return;
        }
        const topSpend = rows[0];
        const bestCvr = [...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const bestVisitorMom = [...rows].sort((a,b)=>safeNum(b.visitorsMom)-safeNum(a.visitorsMom))[0];
        const bestCvrMom = [...rows].sort((a,b)=>safeNum(b.cvrMom)-safeNum(a.cvrMom))[0];
        holder.innerHTML=`<p>${topSpend.label} 是当前花费最高的分组，站内外合计花费 ${fmtMoney(topSpend.innerSpend+topSpend.outerSpend)}，其中站外 ${fmtMoney(topSpend.outerSpend)}、站内 ${fmtMoney(topSpend.innerSpend)}。</p><p>${bestCvr.label} 当前 CVR 最高，为 ${fmtPct(bestCvr.cvr)}；${bestVisitorMom.label} 的访客环比 ${fmtDelta(bestVisitorMom.visitorsMom)}、${bestCvrMom.label} 的 CVR 环比 ${fmtDelta(bestCvrMom.cvrMom)}，更适合用来判断增长到底先来自流量回升还是承接改善。</p><p><strong>下一步：</strong></p>${narrativeList([`${topSpend.label} 继续做花费结构复盘，避免站内外投放结构失衡而拉高成本。`,`优先复盘访客环比和 CVR 环比改善最明显的分组，确认是否值得继续加预算。`])}`;
      }
      const previousRenderAllFinal = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllFinal === "function") previousRenderAllFinal(reset);
        redrawLeftGroupChartFinal();
        renderLeftGroupNarrativeFinal();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(function(){ redrawLeftGroupChartFinal(); renderLeftGroupNarrativeFinal(); },0); });
      }else{
        setTimeout(function(){ redrawLeftGroupChartFinal(); renderLeftGroupNarrativeFinal(); },0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      function redrawLeftGroupFinalV2(){
        if(typeof buildDimensionRows !== "function" || typeof renderPlot !== "function" || typeof aggregateBy !== "function" || typeof filterRows !== "function" || typeof indexBy !== "function") return;
        const baseRows = buildDimensionRows("group").filter(r=>r && r.label !== "其他").slice(0,Math.max(1,state.topN));
        if(!baseRows.length) return;
        const priorShop = aggregateBy(filterRows(DATA.shop_daily,"prior"),["group"],["visitors","orders"]);
        const priorMap = indexBy(priorShop,"group");
        const rows = baseRows.map(r=>{
          const prior = priorMap.get(r.label) || {};
          const priorVisitors = safeNum(prior.visitors);
          const priorOrders = safeNum(prior.orders);
          return {...r, priorVisitors, priorCvr: ratio(priorOrders, priorVisitors)};
        });
        const labels = rows.map(r=>shortLabel(r.label,12));
        const card = document.getElementById("group-chart")?.closest(".card");
        const muted = card?.querySelector(".muted");
        if(muted) muted.textContent = "柱：站内外花费；线：访客、转化率及上月同期";
        renderPlot("group-chart",[
          {type:"bar",name:"站外花费",x:labels,y:rows.map(r=>safeNum(r.outerSpend)),customdata:rows.map(r=>r.label),marker:{color:"#4f8df7"},hovertemplate:"<b>%{customdata}</b><br>站外花费：%{y:,}<extra></extra>"},
          {type:"bar",name:"站内花费",x:labels,y:rows.map(r=>safeNum(r.innerSpend)),customdata:rows.map(r=>r.label),marker:{color:"#ffb84d"},hovertemplate:"<b>%{customdata}</b><br>站内花费：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:rows.map(r=>safeNum(r.visitors)),yaxis:"y2",customdata:rows.map(r=>r.label),line:{color:"#1d9bf0",width:3},marker:{size:8,color:"#1d9bf0"},hovertemplate:"<b>%{customdata}</b><br>访客：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"上月同期访客",x:labels,y:rows.map(r=>safeNum(r.priorVisitors)),yaxis:"y2",customdata:rows.map(r=>r.label),line:{color:"#7cc7ff",width:3,dash:"dot"},marker:{size:8,color:"#7cc7ff"},hovertemplate:"<b>%{customdata}</b><br>上月同期访客：%{y:,}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"CVR",x:labels,y:rows.map(r=>safeNum(r.cvr)),yaxis:"y3",customdata:rows.map(r=>r.label),line:{color:"#16a34a",width:3},marker:{size:8,color:"#16a34a"},hovertemplate:"<b>%{customdata}</b><br>CVR：%{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"上月同期CVR",x:labels,y:rows.map(r=>safeNum(r.priorCvr)),yaxis:"y3",customdata:rows.map(r=>r.label),line:{color:"#86efac",width:3,dash:"dot"},marker:{size:8,color:"#86efac"},hovertemplate:"<b>%{customdata}</b><br>上月同期CVR：%{y:.1%}<extra></extra>"}
        ],{
          height:460,
          barmode:"group",
          bargap:.18,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:88,r:92,t:118,b:112},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{...chartLegendLayout(),orientation:"h",x:0,y:1.18,xanchor:"left",yanchor:"bottom",bgcolor:"rgba(255,255,255,.92)"},
          xaxis:{title:chartAxisTitle("分组"),domain:[0,0.93],tickangle:-18,automargin:true,showgrid:false,linecolor:"#dbe3ef",tickfont:{color:"#64748b"}},
          yaxis:{title:chartAxisTitle("花费"),automargin:true,gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:{text:"访客",standoff:8,font:{family:basePlotFont().family,size:12,color:"#1d9bf0"}},overlaying:"y",anchor:"free",side:"right",position:.985,showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#1d9bf0",size:11}},
          yaxis3:{overlaying:"y",anchor:"free",side:"right",position:1,showgrid:false,zeroline:false,showline:false,showticklabels:false,ticks:"",title:{text:""},tickformat:".0%"}
        },basePlotConfig());
      }
      function renderLeftGroupNarrativeV2(){
        if(typeof buildDimensionRows !== "function" || typeof aggregateBy !== "function" || typeof filterRows !== "function" || typeof indexBy !== "function") return;
        const holder = document.getElementById("group-narrative");
        if(!holder) return;
        const baseRows = buildDimensionRows("group").filter(r=>r && r.label !== "其他");
        const priorShop = aggregateBy(filterRows(DATA.shop_daily,"prior"),["group"],["visitors","orders"]);
        const priorMap = indexBy(priorShop,"group");
        const rows = baseRows.map(r=>{
          const prior = priorMap.get(r.label) || {};
          const priorVisitors = safeNum(prior.visitors);
          const priorOrders = safeNum(prior.orders);
          return {...r, priorVisitors, priorCvr: ratio(priorOrders, priorVisitors)};
        }).sort((a,b)=>safeNum((b.outerSpend+b.innerSpend)) - safeNum((a.outerSpend+a.innerSpend)));
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无分组花费与承接数据。</div>';
          return;
        }
        const topSpend = rows[0];
        const bestCvr = [...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const bestVisitorLift = [...rows].sort((a,b)=>pctChange(safeNum(b.visitors),safeNum(b.priorVisitors))-pctChange(safeNum(a.visitors),safeNum(a.priorVisitors)))[0];
        const bestCvrLift = [...rows].sort((a,b)=>pctChange(safeNum(b.cvr),safeNum(b.priorCvr))-pctChange(safeNum(a.cvr),safeNum(a.priorCvr)))[0];
        holder.innerHTML=`<p>${topSpend.label} 是当前花费最高的分组，站内外合计花费 ${fmtMoney(topSpend.innerSpend+topSpend.outerSpend)}，其中站外 ${fmtMoney(topSpend.outerSpend)}、站内 ${fmtMoney(topSpend.innerSpend)}。</p><p>${bestCvr.label} 当前 CVR 最高，为 ${fmtPct(bestCvr.cvr)}；${bestVisitorLift.label} 的访客较上月同期 ${fmtDelta(pctChange(bestVisitorLift.visitors,bestVisitorLift.priorVisitors))}，${bestCvrLift.label} 的 CVR 较上月同期 ${fmtDelta(pctChange(bestCvrLift.cvr,bestCvrLift.priorCvr))}。</p><p><strong>下一步：</strong></p>${narrativeList([`${topSpend.label} 继续做花费结构复盘，避免站内外投放结构失衡而拉高成本。`,`优先复盘上月同期访客和 CVR 改善最明显的分组，确认增长是来自流量回升还是承接优化。`])}`;
      }
      const previousRenderAllV2 = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllV2 === "function") previousRenderAllV2(reset);
        redrawLeftGroupFinalV2();
        renderLeftGroupNarrativeV2();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(function(){ redrawLeftGroupFinalV2(); renderLeftGroupNarrativeV2(); },0); });
      }else{
        setTimeout(function(){ redrawLeftGroupFinalV2(); renderLeftGroupNarrativeV2(); },0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .chart-grid>.card.chart-grid-single-fill{grid-column:1 / -1;display:grid;grid-template-columns:minmax(0,1.2fr) minmax(300px,.8fr);grid-template-areas:"head head" "plot narrative";gap:18px 22px;align-items:start}
    .chart-grid>.card.chart-grid-single-fill>.chart-head{grid-area:head;margin-bottom:0}
    .chart-grid>.card.chart-grid-single-fill>.plot,
    .chart-grid>.card.chart-grid-single-fill>canvas{grid-area:plot;min-width:0;margin-top:0}
    .chart-grid>.card.chart-grid-single-fill>.narrative{grid-area:narrative;margin-top:0;align-self:stretch;height:100%;min-height:100%;padding:22px 24px}
    .chart-grid>.card.chart-grid-single-fill>.narrative p{font-size:14px;line-height:1.95}
    .chart-grid>.card.chart-grid-single-fill>.narrative ul{padding-left:20px}
    @media (max-width:1200px){
      .chart-grid>.card.chart-grid-single-fill{grid-template-columns:1fr;grid-template-areas:"head" "plot" "narrative"}
      .chart-grid>.card.chart-grid-single-fill>.narrative{height:auto;min-height:0}
    }
  </style>
  <script>
    (function(){
      function applySingleChartRowFill(){
        document.querySelectorAll(".chart-grid>.card.chart-grid-single-fill").forEach(card=>card.classList.remove("chart-grid-single-fill"));
        document.querySelectorAll(".chart-grid").forEach(grid=>{
          const cards = Array.from(grid.children || []).filter(el=>el.classList && el.classList.contains("card"));
          if(!cards.length || cards.length % 2 === 0) return;
          const card = cards[cards.length - 1];
          const children = Array.from(card.children || []);
          const hasHead = children.some(el=>el.classList && el.classList.contains("chart-head"));
          const charts = children.filter(el=>el.matches && el.matches(".plot, canvas"));
          const narratives = children.filter(el=>el.classList && el.classList.contains("narrative"));
          if(!hasHead || charts.length !== 1 || narratives.length !== 1) return;
          card.classList.add("chart-grid-single-fill");
        });
        if(window.Plotly && window.Plotly.Plots && typeof window.Plotly.Plots.resize === "function"){
          document.querySelectorAll(".chart-grid>.card.chart-grid-single-fill > .plot").forEach(el=>{
            try{ window.Plotly.Plots.resize(el); }catch(err){}
          });
        }
      }
      const previousRenderAllSingleChartLayout = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllSingleChartLayout === "function") previousRenderAllSingleChartLayout(reset);
        setTimeout(applySingleChartRowFill,0);
      };
      let singleChartResizeTimer = null;
      window.addEventListener("resize", function(){
        clearTimeout(singleChartResizeTimer);
        singleChartResizeTimer = setTimeout(applySingleChartRowFill,120);
      });
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(applySingleChartRowFill,0); });
      }else{
        setTimeout(applySingleChartRowFill,0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .summary-card .delta.up{color:#ef4444;font-weight:800}
    .summary-card .delta.down{color:#16a34a;font-weight:800}
    .summary-card .delta.flat{color:#7f92af;font-weight:700}
    .table-wrap .mom-value{font-weight:800}
    .table-wrap .mom-value.mom-pos{color:#ef4444}
    .table-wrap .mom-value.mom-neg{color:#16a34a}
    .table-wrap .mom-value.mom-flat{color:#64748b}
  </style>
  <script>
    (function(){
      function isMomKey(key){
        return /Mom$/.test(String(key||""));
      }
      function momClass(value){
        const n = safeNum(value);
        if(n > 0.005) return "mom-pos";
        if(n < -0.005) return "mom-neg";
        return "mom-flat";
      }
      function repaintTableMomCells(kind, rows){
        const holder = document.getElementById(`${kind}-table`);
        if(!holder || !state || !state.tableState || !TABLE_COLUMNS[kind]) return;
        const visible = TABLE_COLUMNS[kind].filter(([key])=>state.tableState[kind].visible[key]);
        const list = rows
          .filter(r=>!state.tableState[kind].hiddenRows.has(r.label))
          .sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv))
          .slice(0,state.topN);
        const trs = holder.querySelectorAll("tbody tr");
        trs.forEach((tr, rowIndex)=>{
          const row = list[rowIndex];
          if(!row) return;
          visible.forEach(([key], colIndex)=>{
            if(!isMomKey(key)) return;
            const td = tr.children[colIndex];
            if(!td) return;
            td.innerHTML = `<span class="mom-value ${momClass(row[key])}">${formatCell(key,row[key])}</span>`;
          });
        });
      }
      const previousRenderTableMomColor = window.renderTable;
      window.renderTable = function(kind, rows){
        if(typeof previousRenderTableMomColor === "function") previousRenderTableMomColor(kind, rows);
        repaintTableMomCells(kind, rows || []);
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          if(typeof buildDiagnosisRows === "function"){
            repaintTableMomCells("category", buildDiagnosisRows("category"));
            repaintTableMomCells("product", buildDiagnosisRows("product"));
          }
        });
      }else if(typeof buildDiagnosisRows === "function"){
        repaintTableMomCells("category", buildDiagnosisRows("category"));
        repaintTableMomCells("product", buildDiagnosisRows("product"));
      }
    })();
  </script>
</body>
""",
        1,
    )
    theme_replacements = [
        (
            ':root{--bg:#f5f7fb;--card:#ffffff;--line:#dbe3ef;--text:#1f2937;--muted:#64748b;--blue:#3b82f6;--blue-deep:#1d4ed8;--blue-soft:rgba(59,130,246,.10);--blue-soft-2:#f8fbff;--warn:#dc2626;--shadow:0 10px 30px rgba(15,23,42,.06)}',
            ':root{--bg:#f3f8ff;--card:#ffffff;--line:#d8e3f4;--text:#233456;--muted:#5f7392;--blue:#2E5AA7;--blue-deep:#244883;--blue-soft:rgba(136,197,255,.18);--blue-soft-2:#f4f9ff;--warn:#dc2626;--accent:#FFA62B;--sky:#88C5FF;--cream:#F8E6A0;--shadow:0 14px 34px rgba(46,90,167,.08)}',
        ),
        (
            '.hero{padding:28px;margin-bottom:18px;background:radial-gradient(circle at top right,rgba(59,130,246,.10),transparent 32%),linear-gradient(135deg,rgba(255,255,255,.98),rgba(248,251,255,.98))}',
            '.hero{padding:28px;margin-bottom:18px;background:radial-gradient(circle at top right,rgba(136,197,255,.22),transparent 34%),radial-gradient(circle at left bottom,rgba(46,90,167,.12),transparent 44%),radial-gradient(circle at 78% 18%,rgba(255,166,43,.10),transparent 18%),linear-gradient(135deg,rgba(255,255,255,.98),rgba(243,248,255,.98))}',
        ),
        (
            '.filter-panel{padding:22px 24px;margin-bottom:20px;background:#eef2f8;color:var(--text);position:static}',
            '.filter-panel{padding:22px 24px;margin-bottom:20px;background:#edf5ff;color:var(--text);position:static}',
        ),
        (
            '.chart-head h2{margin:10px 0 10px;font-size:clamp(22px,2.4vw,30px);line-height:1.22;color:#183764}',
            '.chart-head h2{margin:10px 0 10px;font-size:clamp(22px,2.4vw,30px);line-height:1.22;color:#2E5AA7}',
        ),
        (
            '.btn-primary,.mini-btn{border:none;background:linear-gradient(135deg,var(--blue),#6d95ff);color:#fff;border-radius:999px;padding:10px 16px;font-weight:700}',
            '.btn-primary,.mini-btn{border:none;background:linear-gradient(135deg,var(--blue),var(--sky));color:#fff;border-radius:999px;padding:10px 16px;font-weight:700}',
        ),
        (
            '.summary-card{position:relative;padding:14px 16px 13px;min-height:118px;border-radius:20px;background:#fff;border:1px solid rgba(59,130,246,.10);box-shadow:var(--shadow)}',
            '.summary-card{position:relative;padding:14px 16px 13px;min-height:118px;border-radius:20px;background:#ffffff;border:1px solid rgba(46,90,167,.12);box-shadow:var(--shadow)}',
        ),
        (
            '.plot,canvas{display:block;width:100%;min-height:360px;border-radius:18px;background:linear-gradient(180deg,rgba(47,107,255,.04),rgba(255,255,255,.96)),#fff;border:1px solid rgba(47,107,255,.10);overflow:hidden}',
            '.plot,canvas{display:block;width:100%;min-height:360px;border-radius:18px;background:#ffffff;border:1px solid rgba(46,90,167,.10);overflow:hidden}',
        ),
        (
            '.chart-grid .narrative{margin-top:0;padding:20px 22px;border-radius:20px;background:#fbfcff;border:1px solid var(--line)}',
            '.chart-grid .narrative{margin-top:0;padding:20px 22px;border-radius:20px;background:#f6fbff;border:1px solid var(--line)}',
        ),
        ('#183764', '#2E5AA7'),
        ('#16345f', '#244883'),
        ('#1b3b6c', '#244883'),
        ('#6f84a3', '#5f7392'),
        ('#7f92af', '#7286a7'),
        ('#93c5fd', '#FFA62B'),
        ('#dbeafe', '#F8E6A0'),
        ('#8db9ff', '#88C5FF'),
        ('#cfe1ff', '#F8E6A0'),
        ('#60a5fa', '#2E5AA7'),
        ('#4f8df7', '#2E5AA7'),
        ('#ffb84d', '#FFA62B'),
        ('#2f6bff', '#2E5AA7'),
        ('#84b5ff', '#88C5FF'),
        ('#274472', '#2E5AA7'),
        ('#1e3557', '#244883'),
        ('#1d9bf0', '#88C5FF'),
        ('#0ea5e9', '#88C5FF'),
        ('#8b5cf6', '#F2C35B'),
        ('#94a3b8', '#E9D487'),
        ('#7c3aed', '#E8B849'),
        ('#7cc7ff', '#B8DCFF'),
        ('#86efac', '#F4D87C'),
        ('#f1f6ff', '#edf5ff'),
        ('#fbfcff', '#f6fbff'),
        ('rgba(47,107,255', 'rgba(46,90,167'),
        ('rgba(59,130,246', 'rgba(136,197,255'),
    ]
    for old, new in theme_replacements:
        html = html.replace(old, new)
    html = html.replace(
        '.summary-card .delta.up{color:#FFA62B;font-weight:800}',
        '.summary-card .delta.up{color:#ef4444;font-weight:800}',
    )
    html = html.replace(
        '.summary-card .delta.down{color:#FFA62B;font-weight:800}',
        '.summary-card .delta.down{color:#16a34a;font-weight:800}',
    )
    html = html.replace(
        '.table-wrap .mom-value.mom-pos{color:#FFA62B}',
        '.table-wrap .mom-value.mom-pos{color:#ef4444}',
    )
    html = html.replace(
        '.table-wrap .mom-value.mom-neg{color:#FFA62B}',
        '.table-wrap .mom-value.mom-neg{color:#16a34a}',
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .filter-panel,
    .summary-card,
    .plot,
    canvas,
    .chart-grid .narrative{
      background-image:none !important;
    }
    .filter-panel{background:#edf5ff !important}
    .summary-card{background:#ffffff !important}
    .plot,canvas{background:#ffffff !important}
    .chart-grid .narrative{background:#f6fbff !important}
  </style>
  <script>
    (function(){
      function priorMetricField(field){
        const map={
          gmv:"priorGmv",
          visitors:"priorVisitors",
          cvr:"priorCvr",
          innerSpend:"priorInnerSpend",
          outerSpend:"priorOuterSpend",
          spend:"priorSpend"
        };
        return map[field] || "";
      }
      function buildDimensionRowsWithPrior(dim){
        const cs=aggregateBy(filterRows(DATA.shop_daily,"current"),[dim],["gmv","visitors","orders"]);
        const ps=aggregateBy(filterRows(DATA.shop_daily,"prior"),[dim],["gmv","visitors","orders"]);
        const co=aggregateBy(filterRows(DATA.outer_daily,"current"),[dim],["outer_spend","impressions","clicks"]);
        const po=aggregateBy(filterRows(DATA.outer_daily,"prior"),[dim],["outer_spend","impressions","clicks"]);
        const ci=aggregateBy(filterRows(DATA.inner_daily,"current"),[dim],["inner_spend"]);
        const pi=aggregateBy(filterRows(DATA.inner_daily,"prior"),[dim],["inner_spend"]);
        const csMap=indexBy(cs,dim),psMap=indexBy(ps,dim),coMap=indexBy(co,dim),poMap=indexBy(po,dim),ciMap=indexBy(ci,dim),piMap=indexBy(pi,dim);
        const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);
        const rows=Array.from(keys).map(key=>{
          const current={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};
          const prior={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};
          const g=safeNum(current.gmv),v=safeNum(current.visitors),o=safeNum(current.orders);
          const pg=safeNum(prior.gmv),pv=safeNum(prior.visitors),poo=safeNum(prior.orders);
          const cis=safeNum(current.inner_spend),pis=safeNum(prior.inner_spend),cos=safeNum(current.outer_spend),pos=safeNum(prior.outer_spend);
          const imp=safeNum(current.impressions),pimp=safeNum(prior.impressions);
          return {
            label:key,
            gmv:g,
            priorGmv:pg,
            visitors:v,
            priorVisitors:pv,
            orders:o,
            priorOrders:poo,
            cvr:ratio(o,v),
            priorCvr:ratio(poo,pv),
            innerSpend:cis,
            priorInnerSpend:pis,
            outerSpend:cos,
            priorOuterSpend:pos,
            impressions:imp,
            priorImpressions:pimp,
            gmvMom:pctChange(g,pg),
            ordersMom:pctChange(o,poo),
            visitorsMom:pctChange(v,pv),
            cvrMom:pctChange(ratio(o,v),ratio(poo,pv)),
            innerSpendMom:pctChange(cis,pis),
            outerSpendMom:pctChange(cos,pos),
            impressionsMom:pctChange(imp,pimp)
          };
        }).filter(r=>r.label&&r.label!=="未归类"&&r.label!=="鏈綊绫?"&&r.label!=="閺堫亜缍婄猾?"&&r.label!=="闁哄牜浜滅紞濠勭尵?");
        const cfg=getMetricConfig();
        return rows.sort((a,b)=>safeNum(b[cfg.field])-safeNum(a[cfg.field]));
      }
      function buildModuleRowsWithPrior(){
        const current=aggregateBy(filterRows(DATA.outer_module_daily,"current"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const prior=aggregateBy(filterRows(DATA.outer_module_daily,"prior"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const currentMap=indexBy(current,"module"), priorMap=indexBy(prior,"module");
        const keys=new Set([...currentMap.keys(),...priorMap.keys()]);
        return Array.from(keys).map(key=>{
          const r=currentMap.get(key)||{};
          const pr=priorMap.get(key)||{};
          return {
            label:key,
            spend:safeNum(r.outer_spend),
            priorSpend:safeNum(pr.outer_spend),
            impressions:safeNum(r.impressions),
            priorImpressions:safeNum(pr.impressions),
            clicks:safeNum(r.clicks),
            priorClicks:safeNum(pr.clicks),
            ctr:ratio(r.clicks,r.impressions),
            priorCtr:ratio(pr.clicks,pr.impressions),
            roi:ratio(r.ad_gmv,r.outer_spend),
            priorRoi:ratio(pr.ad_gmv,pr.outer_spend),
            spendMom:pctChange(safeNum(r.outer_spend),safeNum(pr.outer_spend))
          };
        }).filter(r=>r.label&&r.label!=="未归类"&&r.label!=="鏈綊绫?").sort((a,b)=>safeNum(b.spend)-safeNum(a.spend));
      }
      function drawBarsWithPrior(id,rows,labelField,valueField,deltaField,formatter){
        const data=rows.slice(0,state.topN);
        const priorField=priorMetricField(valueField);
        const isPct=valueField==="cvr";
        const labels=data.map(r=>shortLabel(r[labelField],12));
        renderPlot(id,[
          {
            type:"bar",
            name:`4月${chartMetricLabel(valueField)}`,
            x:labels,
            y:data.map(r=>safeNum(r[valueField])),
            marker:{color:"#FFA62B"},
            text:data.map(r=>formatter(r[valueField])),
            textposition:"outside",
            cliponaxis:false,
            customdata:data.map(r=>[r[labelField],safeNum(r[deltaField]),priorField?safeNum(r[priorField]):0]),
            hovertemplate:`<b>%{customdata[0]}</b><br>${chartMetricLabel(valueField)}: ${isPct?"%{y:.1%}":"%{y:,}"}<br>MoM: %{customdata[1]:+.1%}<extra></extra>`
          },
          {
            type:"bar",
            name:"3月同期",
            x:labels,
            y:data.map(r=>priorField?safeNum(r[priorField]):0),
            marker:{color:"#F8E6A0"},
            text:data.map(r=>priorField?formatter(r[priorField]):formatter(0)),
            textposition:"outside",
            cliponaxis:false,
            hovertemplate:`<b>%{x}</b><br>3月同期${chartMetricLabel(valueField)}: ${isPct?"%{y:.1%}":"%{y:,}"}<extra></extra>`
          }
        ],{
          height:Math.max(430,360+Math.ceil(data.length/12)*24),
          barmode:"group",
          bargap:.2,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:84,r:28,t:28,b:132},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{...chartLegendLayout(),orientation:"h",x:1,y:1.12},
          xaxis:{tickangle:-28,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:chartAxisTitle(chartMetricLabel(valueField)),gridcolor:"#eef3fb",zeroline:false,tickformat:isPct?".0%":",.0f",tickfont:{color:"#64748b"},automargin:true}
        },basePlotConfig());
      }
      buildDimensionRows = buildDimensionRowsWithPrior;
      buildModuleRows = buildModuleRowsWithPrior;
      drawBars = drawBarsWithPrior;
      buildBundleRows = function(){
        const productRows=buildDiagnosisRows("product");
        const map=indexBy(productRows,"label");
        return DATA.focus_bundles.map(name=>map.get(name)||{
          label:name,gmv:0,priorGmv:0,visitors:0,priorVisitors:0,cvr:0,priorCvr:0,innerSpend:0,priorInnerSpend:0,outerSpend:0,priorOuterSpend:0,impressions:0,priorImpressions:0,gmvMom:0,visitorsMom:0,cvrMom:0,
          diagnosis:"当前筛选范围内暂无有效店铺或投放数据。",todo:"先确认映射与素材是否已上线，再决定是否放量。"
        });
      };
      setTimeout(function(){
        if(typeof renderAll==="function") renderAll(true);
      },0);
    })();
  </script>
  <script>
    (function(){
      function setCardTitle(chartId, nextTitle){
        const chart = document.getElementById(chartId);
        const card = chart && chart.closest(".card");
        const title = card && card.querySelector(".chart-head h2");
        if(title) title.textContent = nextTitle;
      }
      function swapGroupCards(){
        const groupChart = document.getElementById("group-chart");
        const compareChart = document.getElementById("group-compare-chart");
        const groupCard = groupChart && groupChart.closest(".card");
        const compareCard = compareChart && compareChart.closest(".card");
        const grid = groupCard && groupCard.parentElement;
        if(!grid || !compareCard || groupCard.parentElement !== compareCard.parentElement) return;
        if(compareCard.nextElementSibling === groupCard) return;
        grid.insertBefore(compareCard, groupCard);
      }
      function applyTitleAndOrderTweaks(){
        setCardTitle("gmv-trend","GMV日趋");
        setCardTitle("spend-mix","站内外花费与店铺GMV");
        setCardTitle("group-compare-chart","分组GMV");
        swapGroupCards();
      }
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(applyTitleAndOrderTweaks, 0); });
      }else{
        setTimeout(applyTitleAndOrderTweaks, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .shop-analysis-card{background:#ffffff !important}
    .shop-stack{display:grid;gap:18px}
    .shop-block{padding:18px;border:1px solid rgba(47,107,255,.12);border-radius:22px;background:#f8fbff}
    .shop-block-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px}
    .shop-tag{display:inline-flex;align-items:center;min-height:38px;padding:8px 16px;border-radius:12px;background:#4d98ff;color:#ffffff;font-size:14px;font-weight:700;line-height:1.5}
    .shop-mini-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}
    .shop-rank-wrap{padding:18px;border:1px solid rgba(47,107,255,.12);border-radius:22px;background:#f8fbff}
    .shop-plot{width:100%;background:#ffffff;border-radius:18px}
    .shop-plot-medium{min-height:340px}
    .shop-plot-small{min-height:240px}
    .shop-table-wrap{overflow:auto;border-radius:18px;background:#ffffff;border:1px solid rgba(148,163,184,.14)}
    .shop-table-wrap table{width:100%;min-width:920px;border-collapse:separate;border-spacing:0}
    .shop-table-wrap th,.shop-table-wrap td{padding:12px 14px;border-bottom:1px solid #edf2f9;text-align:left;font-size:13px;white-space:nowrap}
    .shop-table-wrap th{position:sticky;top:0;background:#f5f9ff;color:#31507e;font-weight:700;z-index:1}
    .shop-table-wrap tr:last-child td{border-bottom:none}
    .shop-mom{font-weight:800}
    .shop-mom.is-up{color:#cf2e2e}
    .shop-mom.is-down{color:#179c53}
    .shop-mom.is-flat{color:#64748b}
    .shop-rank-list{margin:14px 0 0;padding:0;list-style:none;display:grid;gap:10px}
    .shop-rank-item{display:grid;grid-template-columns:42px minmax(0,1fr) auto;gap:10px;align-items:center;padding:10px 12px;border-radius:14px;background:#ffffff;border:1px solid rgba(148,163,184,.14)}
    .shop-rank-no{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:10px;background:#e9f2ff;color:#2f6bff;font-weight:800}
    .shop-rank-name{min-width:0;color:#1e293b;font-weight:700;line-height:1.45}
    .shop-rank-metric{color:#5b6b82;font-size:12px;line-height:1.45;text-align:right}
    .shop-rank-metric strong{display:block;color:#2f6bff;font-size:13px}
    .shop-mini-empty{padding:18px 12px;color:#7b8ba5;font-size:13px;line-height:1.7}
    .shop-analysis-card .chart-tools button{white-space:nowrap}
    @media (max-width: 960px){
      .shop-mini-grid{grid-template-columns:1fr}
      .shop-table-wrap table{min-width:760px}
    }
  </style>
  <script>
    (function(){
      const SHOP_COLORS=["#2f6bff","#67a6ff","#9ec6ff","#7dd3fc","#dbeafe"];
      function shopMomClass(value){
        const n=safeNum(value);
        if(n>0.005) return "is-up";
        if(n<-0.005) return "is-down";
        return "is-flat";
      }
      function shopMomHtml(value){
        return `<span class="shop-mom ${shopMomClass(value)}">${fmtDelta(value)}</span>`;
      }
      function insertShopAnalysisCards(){
        if(document.getElementById("shop-share-chart")) return;
        const spendChart=document.getElementById("spend-mix");
        const spendCard=spendChart&&spendChart.closest(".card");
        const grid=spendCard&&spendCard.parentElement;
        if(!grid||!spendCard) return;
        const wrap=document.createElement("div");
        wrap.innerHTML=`
          <div class="card shop-analysis-card" id="shop-analysis-left-card">
            <div class="chart-head">
              <div>
                <div class="section-label">Shop Analysis</div>
                <h2>店铺分析</h2>
                <div class="muted">聚焦各店铺贡献结构与核心经营诊断</div>
              </div>
              <div class="chart-tools">
                <button type="button" data-shop-download="shop-share-chart">下载图表</button>
              </div>
            </div>
            <div class="shop-stack">
              <div class="shop-block">
                <div class="shop-tag">饼图：显示各店铺GMV占比</div>
                <div id="shop-share-chart" class="plot shop-plot shop-plot-medium"></div>
              </div>
              <div class="shop-block">
                <div class="shop-tag">表格：分店铺的GMV、访客、转化率、站内花费、站外花费，以及相应的环比</div>
                <div id="shop-table-holder" class="shop-table-wrap"></div>
              </div>
            </div>
            <div class="narrative" id="shop-left-narrative"></div>
          </div>
          <div class="card shop-analysis-card" id="shop-analysis-right-card">
            <div class="chart-head">
              <div>
                <div class="section-label">Shop Conversion</div>
                <h2>店铺承接分析</h2>
                <div class="muted">看店铺转化、花费承接和主力产品结构</div>
              </div>
            </div>
            <div class="shop-stack">
              <div class="shop-block">
                <div class="shop-block-head">
                  <div class="shop-tag">曲线图：转化率，本月和上月同期</div>
                  <button type="button" data-shop-download="shop-cvr-chart">下载图表</button>
                </div>
                <div id="shop-cvr-chart" class="plot shop-plot shop-plot-small"></div>
              </div>
              <div class="shop-block">
                <div class="shop-block-head">
                  <div class="shop-tag">组合图：站外花费柱+访客曲线，本月和上月同期</div>
                  <button type="button" data-shop-download="shop-spend-visitor-chart">下载图表</button>
                </div>
                <div id="shop-spend-visitor-chart" class="plot shop-plot shop-plot-medium"></div>
              </div>
              <div class="shop-mini-grid">
                <div class="shop-rank-wrap">
                  <div class="shop-tag" id="shop-top-a-title-legacy">店铺GMV Top10产品</div>
                  <div id="shop-top-a-legacy"></div>
                </div>
                <div class="shop-rank-wrap">
                  <div class="shop-tag" id="shop-top-b-title-legacy">店铺GMV Top10产品</div>
                  <div id="shop-top-b-legacy"></div>
                </div>
              </div>
            </div>
            <div class="narrative" id="shop-right-narrative-legacy"></div>
          </div>`;
        const nodes=Array.from(wrap.children);
        if(!nodes.length) return;
        grid.insertBefore(nodes[0], spendCard.nextSibling);
        if(nodes[1]) grid.insertBefore(nodes[1], nodes[0].nextSibling);
        bindShopDownloads();
      }
      function bindShopDownloads(){
        document.querySelectorAll("[data-shop-download]").forEach(btn=>{
          btn.onclick=()=>downloadCanvas(btn.dataset.shopDownload);
        });
      }
      function buildShopRows(){
        if(typeof buildDimensionRows!=="function") return [];
        return buildDimensionRows("shop").filter(row=>row&&row.label&&row.label!=="其他");
      }
      function renderShopShare(rows){
        const holder=document.getElementById("shop-share-chart");
        if(!holder) return;
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺占比数据。</div>';
          return;
        }
        renderPlot("shop-share-chart",[
          {
            type:"pie",
            labels:rows.map(r=>r.label),
            values:rows.map(r=>safeNum(r.gmv)),
            hole:.58,
            sort:false,
            direction:"clockwise",
            textinfo:"percent",
            textposition:"inside",
            insidetextorientation:"radial",
            marker:{colors:rows.map((_,idx)=>SHOP_COLORS[idx%SHOP_COLORS.length]),line:{width:0}},
            hovertemplate:"<b>%{label}</b><br>GMV: %{value:,.0f}<br>占比: %{percent}<extra></extra>"
          }
        ],{
          height:340,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:8,r:128,t:10,b:10},
          font:basePlotFont(),
          showlegend:true,
          legend:{orientation:"v",x:1.02,xanchor:"left",y:.5,yanchor:"middle",bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1,font:{family:basePlotFont().family,size:12,color:"#334155"}}
        },basePlotConfig());
      }
      function renderShopTable(rows){
        const holder=document.getElementById("shop-table-holder");
        if(!holder) return;
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺诊断表数据。</div>';
          return;
        }
        holder.innerHTML=`
          <table>
            <thead>
              <tr>
                <th>店铺</th>
                <th>GMV</th>
                <th>环比</th>
                <th>访客</th>
                <th>环比</th>
                <th>转化率</th>
                <th>环比</th>
                <th>站内花费</th>
                <th>环比</th>
                <th>站外花费</th>
                <th>环比</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(row=>`
                <tr>
                  <td>${row.label}</td>
                  <td>${fmtMoney(row.gmv)}</td>
                  <td>${shopMomHtml(row.gmvMom)}</td>
                  <td>${fmtNum(row.visitors)}</td>
                  <td>${shopMomHtml(row.visitorsMom)}</td>
                  <td>${fmtPct(row.cvr)}</td>
                  <td>${shopMomHtml(row.cvrMom)}</td>
                  <td>${fmtMoney(row.innerSpend)}</td>
                  <td>${shopMomHtml(row.innerSpendMom)}</td>
                  <td>${fmtMoney(row.outerSpend)}</td>
                  <td>${shopMomHtml(row.outerSpendMom)}</td>
                </tr>`).join("")}
            </tbody>
          </table>`;
      }
      function renderShopCvr(rows){
        const holder=document.getElementById("shop-cvr-chart");
        if(!holder) return;
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺转化率数据。</div>';
          return;
        }
        const labels=rows.map(r=>shortLabel(r.label,12));
        renderPlot("shop-cvr-chart",[
          {
            type:"scatter",
            mode:"lines+markers",
            name:"4月CVR",
            x:labels,
            y:rows.map(r=>safeNum(r.cvr)),
            customdata:rows.map(r=>r.label),
            line:{color:"#2f6bff",width:3},
            marker:{size:8,color:"#2f6bff"},
            hovertemplate:"<b>%{customdata}</b><br>4月CVR: %{y:.1%}<extra></extra>"
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:"3月同期CVR",
            x:labels,
            y:rows.map(r=>safeNum(r.priorCvr)),
            customdata:rows.map(r=>r.label),
            line:{color:"#9ec6ff",width:3},
            marker:{size:8,color:"#9ec6ff"},
            hovertemplate:"<b>%{customdata}</b><br>3月同期CVR: %{y:.1%}<extra></extra>"
          }
        ],{
          height:240,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:60,r:18,t:38,b:84},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-18,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"CVR",tickformat:".0%",gridcolor:"#eef3fb",zeroline:false,tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }
      function renderShopSpendVisitors(rows){
        const holder=document.getElementById("shop-spend-visitor-chart");
        if(!holder) return;
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺承接数据。</div>';
          return;
        }
        const labels=rows.map(r=>shortLabel(r.label,12));
        renderPlot("shop-spend-visitor-chart",[
          {
            type:"bar",
            name:"4月站外花费",
            x:labels,
            y:rows.map(r=>safeNum(r.outerSpend)),
            customdata:rows.map(r=>r.label),
            marker:{color:"#2f6bff",line:{width:0}},
            hovertemplate:"<b>%{customdata}</b><br>4月站外花费: %{y:,.0f}<extra></extra>"
          },
          {
            type:"bar",
            name:"3月同期站外花费",
            x:labels,
            y:rows.map(r=>safeNum(r.priorOuterSpend)),
            customdata:rows.map(r=>r.label),
            marker:{color:"#b7d5ff",line:{width:0}},
            hovertemplate:"<b>%{customdata}</b><br>3月同期站外花费: %{y:,.0f}<extra></extra>"
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:"4月访客",
            x:labels,
            y:rows.map(r=>safeNum(r.visitors)),
            customdata:rows.map(r=>r.label),
            yaxis:"y2",
            line:{color:"#f59e0b",width:3},
            marker:{size:8,color:"#f59e0b"},
            hovertemplate:"<b>%{customdata}</b><br>4月访客: %{y:,.0f}<extra></extra>"
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:"3月同期访客",
            x:labels,
            y:rows.map(r=>safeNum(r.priorVisitors)),
            customdata:rows.map(r=>r.label),
            yaxis:"y2",
            line:{color:"#fcd34d",width:3},
            marker:{size:8,color:"#fcd34d"},
            hovertemplate:"<b>%{customdata}</b><br>3月同期访客: %{y:,.0f}<extra></extra>"
          }
        ],{
          height:340,
          barmode:"group",
          bargap:.22,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:64,r:62,t:46,b:88},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.2,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-18,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"站外花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:"访客",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }
      function buildShopTopProducts(shopName){
        if(!shopName) return [];
        const rows=aggregateBy(filterRows(DATA.shop_daily,"current").filter(row=>row.shop===shopName),["product"],["gmv","visitors","orders"]);
        return rows.map(row=>{
          const g=safeNum(row.gmv),v=safeNum(row.visitors),o=safeNum(row.orders);
          return {label:row.product||"未归类",gmv:g,visitors:v,cvr:ratio(o,v)};
        }).filter(row=>row.label&&row.label!=="未归类").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,10);
      }
      function renderShopTopProducts(holderId,titleId,shopName){
        const holder=document.getElementById(holderId);
        const title=document.getElementById(titleId);
        if(!holder||!title) return;
        title.textContent=shopName?`${shopName}店GMV Top10产品`:"店铺GMV Top10产品";
        const rows=buildShopTopProducts(shopName);
        if(!shopName){
          holder.innerHTML='<div class="shop-mini-empty">当前筛选下暂无可展示的店铺。</div>';
          return;
        }
        if(!rows.length){
          holder.innerHTML='<div class="shop-mini-empty">当前筛选下该店铺暂无有效产品数据。</div>';
          return;
        }
        holder.innerHTML=`<div class="shop-rank-list">${rows.map((row,idx)=>`
          <div class="shop-rank-item">
            <span class="shop-rank-no">${idx+1}</span>
            <div class="shop-rank-name">${row.label}</div>
            <div class="shop-rank-metric"><strong>${fmtMoney(row.gmv)}</strong><span>访客 ${fmtNum(row.visitors)} | CVR ${fmtPct(row.cvr)}</span></div>
          </div>`).join("")}</div>`;
      }
      function renderShopNarratives(rows){
        const left=document.getElementById("shop-left-narrative");
        const right=document.getElementById("shop-right-narrative-legacy");
        if(!left||!right) return;
        if(!rows.length){
          left.innerHTML='<div class="empty-state">当前筛选下暂无店铺分析结论。</div>';
          right.innerHTML='<div class="empty-state">当前筛选下暂无店铺承接分析结论。</div>';
          return;
        }
        const totalGmv=rows.reduce((sum,row)=>sum+safeNum(row.gmv),0);
        const totalVisitors=rows.reduce((sum,row)=>sum+safeNum(row.visitors),0);
        const totalOrders=rows.reduce((sum,row)=>sum+safeNum(row.orders),0);
        const top=rows[0];
        const second=rows[1];
        const bestCvr=[...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const spendLead=[...rows].sort((a,b)=>safeNum(b.outerSpend)-safeNum(a.outerSpend))[0];
        const avgCvr=ratio(totalOrders,totalVisitors);
        const concentration=ratio(top.gmv,totalGmv);
        const focusShops=rows.slice(0,2).map(r=>r.label);
        left.innerHTML=`
          <p>${top.label} 当前店铺GMV ${fmtMoney(top.gmv)}，贡献占比 ${fmtPct(concentration)}；${second?`${second.label} 以 ${fmtMoney(second.gmv)} 排名第二。`:"当前筛选仅保留单店铺样本。"}</p>
          <p>${concentration>=0.45?"当前店铺GMV集中度偏高，核心增长更依赖头部店铺稳定承接。":"当前店铺GMV分布相对均衡，说明不同店铺都具备一定承接能力。"} ${bestCvr.label} 的CVR最高，为 ${fmtPct(bestCvr.cvr)}，较整体均值 ${fmtPct(avgCvr)} ${bestCvr.cvr>=avgCvr?"更强":"略弱"}。</p>
          <p><strong>下一步：</strong></p>
          ${narrativeList([
            `${top.label} 继续稳住主力SKU和重点套组承接，避免头部店铺因素材疲劳带动整体GMV波动。`,
            `${bestCvr.label!==top.label?`复盘 ${bestCvr.label} 的详情页、价格带和评论资产，把高转化承接经验复制给 ${top.label} 等规模店铺。`:`继续保持 ${bestCvr.label} 的转化优势，并适度放大流量验证上限。`}`,
            `${focusShops.join("、")} 之外的店铺若持续低占比，建议明确角色分工，区分为规模盘、测试盘或补量盘。`
          ])}`;
        right.innerHTML=`
          <p>${spendLead.label} 当前站外花费最高，为 ${fmtMoney(spendLead.outerSpend)}，对应访客 ${fmtNum(spendLead.visitors)}，CVR ${fmtPct(spendLead.cvr)}，环比 ${shopMomHtml(spendLead.cvrMom)}。</p>
          <p>${bestCvr.label} 是当前承接效率最好的店铺，若其流量规模仍偏小，可优先从相近人群或相同素材角度做小步放量；${top.label} 作为GMV头部店铺，更需要平衡规模与效率，避免高花费下访客增长但转化走弱。</p>
          <p><strong>下一步：</strong></p>
          ${narrativeList([
            `${spendLead.label} 重点排查高花费素材的人群匹配度和落地页承接，确认访客增长是否真正带来订单提升。`,
            `${bestCvr.label} 可作为高转化店铺模板，对照其高效产品结构继续补量。`,
            `${focusShops.join("、")} 的Top产品建议持续做店铺差异化布局，减少同质SKU在不同店铺之间互相抢量。`
          ])}`;
      }
      function renderShopAnalysis(){
        insertShopAnalysisCards();
        const rows=buildShopRows();
        renderShopShare(rows);
        renderShopTable(rows);
        renderShopCvr(rows);
        renderShopSpendVisitors(rows);
        renderShopTopProducts("shop-top-a-legacy","shop-top-a-title-legacy","Authorize");
        renderShopTopProducts("shop-top-b-legacy","shop-top-b-title-legacy","Certified");
        renderShopNarratives(rows);
      }
      const previousShopRenderAll=window.renderAll;
      window.renderAll=function(reset){
        if(typeof previousShopRenderAll==="function") previousShopRenderAll(reset);
        renderShopAnalysis();
      };
      if(document.readyState==="loading"){
        document.addEventListener("DOMContentLoaded",function(){setTimeout(renderShopAnalysis,0);});
      }else{
        setTimeout(renderShopAnalysis,0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    #shop-analysis-left-card,
    #shop-analysis-right-card{min-width:0;overflow:hidden}
    #shop-analysis-left-card .shop-block,
    #shop-analysis-right-card .shop-block,
    #shop-analysis-right-card .shop-mini-grid,
    #shop-analysis-right-card .shop-rank-wrap{min-width:0}
    #shop-table-holder{width:100%;max-width:100%;max-height:360px;overflow:auto}
    #shop-table-holder table{width:max-content;min-width:1000px;max-width:none}
    #shop-analysis-right-card .shop-mini-grid{align-items:stretch}
    #shop-analysis-right-card .shop-rank-wrap{display:grid;grid-template-rows:auto minmax(0,1fr);height:360px;max-height:360px;overflow:hidden}
    #shop-top-a,#shop-top-b{min-height:0;height:100%;overflow:auto;padding-right:4px}
  </style>
  <script>
    (function(){
      function buildShopDailySeriesFinal(){
        if(typeof aggregateBy!=="function" || typeof filterRows!=="function" || typeof indexBy!=="function") return null;
        const dayCount = Number(DATA?.meta?.day_count || 0);
        const days = Array.from({length:dayCount},(_,i)=>i+1);
        const currentShop = indexBy(aggregateBy(filterRows(DATA.shop_daily,"current"),["day"],["visitors","orders"]),"day");
        const priorShop = indexBy(aggregateBy(filterRows(DATA.shop_daily,"prior"),["day"],["visitors","orders"]),"day");
        const currentOuter = indexBy(aggregateBy(filterRows(DATA.outer_daily,"current"),["day"],["outer_spend"]),"day");
        const priorOuter = indexBy(aggregateBy(filterRows(DATA.outer_daily,"prior"),["day"],["outer_spend"]),"day");
        const monthNum = Number(String(DATA?.meta?.current_start || "").slice(5,7)) || 4;
        return {
          labels: days.map(day => `${monthNum}/${day}`),
          cvrCurrent: days.map(day => { const row=currentShop.get(day)||{}; return ratio(row.orders,row.visitors); }),
          cvrPrior: days.map(day => { const row=priorShop.get(day)||{}; return ratio(row.orders,row.visitors); }),
          visitorsCurrent: days.map(day => safeNum((currentShop.get(day)||{}).visitors)),
          visitorsPrior: days.map(day => safeNum((priorShop.get(day)||{}).visitors)),
          outerCurrent: days.map(day => safeNum((currentOuter.get(day)||{}).outer_spend)),
          outerPrior: days.map(day => safeNum((priorOuter.get(day)||{}).outer_spend))
        };
      }
      function shopTickLabels(labels){
        return labels.filter((_,idx)=>idx===0||idx===labels.length-1||idx%2===1);
      }
      function redrawShopDateChartsFinal(){
        if(typeof renderPlot!=="function") return;
        const cvrHolder = document.getElementById("shop-cvr-chart");
        const spendHolder = document.getElementById("shop-spend-visitor-chart");
        if(!cvrHolder || !spendHolder) return;
        const daily = buildShopDailySeriesFinal();
        if(!daily || !daily.labels.length) return;
        const tickLabels = shopTickLabels(daily.labels);
        const hasCvr = daily.cvrCurrent.concat(daily.cvrPrior).some(v=>safeNum(v)>0);
        const hasSpend = daily.outerCurrent.concat(daily.outerPrior,daily.visitorsCurrent,daily.visitorsPrior).some(v=>safeNum(v)>0);
        if(!hasCvr){
          cvrHolder.innerHTML='<div class="empty-state">当前筛选下暂无店铺转化率数据。</div>';
        }else{
          renderPlot("shop-cvr-chart",[
            {type:"scatter",mode:"lines+markers",name:"4月CVR",x:daily.labels,y:daily.cvrCurrent,line:{color:"#2f6bff",width:3},marker:{size:7,color:"#2f6bff"},hovertemplate:"<b>%{x}</b><br>4月CVR: %{y:.1%}<extra></extra>"},
            {type:"scatter",mode:"lines+markers",name:"3月同期CVR",x:daily.labels,y:daily.cvrPrior,line:{color:"#9ec6ff",width:3},marker:{size:7,color:"#9ec6ff"},hovertemplate:"<b>%{x}</b><br>3月同期CVR: %{y:.1%}<extra></extra>"}
          ],{
            height:240,
            paper_bgcolor:"rgba(0,0,0,0)",
            plot_bgcolor:"#ffffff",
            margin:{l:60,r:18,t:38,b:72},
            hovermode:"x unified",
            hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
            font:basePlotFont(),
            legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
            xaxis:{title:"日期",tickmode:"array",tickvals:tickLabels,ticktext:tickLabels,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
            yaxis:{title:"CVR",tickformat:".0%",gridcolor:"#eef3fb",zeroline:false,tickfont:{color:"#64748b"}}
          },basePlotConfig());
        }
        if(!hasSpend){
          spendHolder.innerHTML='<div class="empty-state">当前筛选下暂无店铺承接数据。</div>';
        }else{
          renderPlot("shop-spend-visitor-chart",[
            {type:"bar",name:"4月站外花费",x:daily.labels,y:daily.outerCurrent,marker:{color:"#2f6bff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>4月站外花费: %{y:,.0f}<extra></extra>"},
            {type:"bar",name:"3月同期站外花费",x:daily.labels,y:daily.outerPrior,marker:{color:"#b7d5ff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>3月同期站外花费: %{y:,.0f}<extra></extra>"},
            {type:"scatter",mode:"lines+markers",name:"4月访客",x:daily.labels,y:daily.visitorsCurrent,yaxis:"y2",line:{color:"#f59e0b",width:3},marker:{size:7,color:"#f59e0b"},hovertemplate:"<b>%{x}</b><br>4月访客: %{y:,.0f}<extra></extra>"},
            {type:"scatter",mode:"lines+markers",name:"3月同期访客",x:daily.labels,y:daily.visitorsPrior,yaxis:"y2",line:{color:"#fcd34d",width:3},marker:{size:7,color:"#fcd34d"},hovertemplate:"<b>%{x}</b><br>3月同期访客: %{y:,.0f}<extra></extra>"}
          ],{
            height:340,
            barmode:"group",
            bargap:.22,
            paper_bgcolor:"rgba(0,0,0,0)",
            plot_bgcolor:"#ffffff",
            margin:{l:64,r:62,t:46,b:72},
            hovermode:"x unified",
            hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
            font:basePlotFont(),
            legend:{orientation:"h",x:0,y:1.2,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
            xaxis:{title:"日期",tickmode:"array",tickvals:tickLabels,ticktext:tickLabels,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
            yaxis:{title:"站外花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
            yaxis2:{title:"访客",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
          },basePlotConfig());
        }
      }
      const prevRenderAllShopDateFinal = window.renderAll;

      window.renderAll = function(reset){
        if(typeof prevRenderAllShopDateFinal === "function") prevRenderAllShopDateFinal(reset);
        redrawShopDateChartsFinal();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(redrawShopDateChartsFinal, 0); });
      }else{
        setTimeout(redrawShopDateChartsFinal, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    #shop-top-products-card{grid-column:2}
    #shop-top-products-card .shop-top-card-body{min-width:0}
    #shop-top-products-card .chart-head{align-items:flex-end}
    .shop-top-filter{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
    .shop-top-filter label{font-size:12px;color:#5f7392;font-weight:700}
    .shop-top-filter select{height:42px;min-width:132px;border-radius:14px;border:1px solid rgba(46,90,167,.16);padding:0 14px;background:#fff;color:#233456;font:inherit}
    @media (max-width:960px){
      #shop-top-products-card{grid-column:auto}
    }
  </style>
  <script>
    (function(){
      function shopTopDayLabel(dayValue){
        const monthNum = Number(String(DATA?.meta?.current_start || "").slice(5,7)) || 4;
        return dayValue ? `${monthNum}/${dayValue}` : "全部日期";
      }
      function buildShopDailySeriesV2(){
        if(typeof aggregateBy!=="function" || typeof filterRows!=="function" || typeof indexBy!=="function") return null;
        const dayCount = Number(DATA?.meta?.day_count || 0);
        const days = Array.from({length:dayCount},(_,i)=>i+1);
        const currentShop = indexBy(aggregateBy(filterRows(DATA.shop_daily,"current"),["day"],["visitors","orders"]),"day");
        const priorShop = indexBy(aggregateBy(filterRows(DATA.shop_daily,"prior"),["day"],["visitors","orders"]),"day");
        const currentOuter = indexBy(aggregateBy(filterRows(DATA.outer_daily,"current"),["day"],["outer_spend"]),"day");
        const priorOuter = indexBy(aggregateBy(filterRows(DATA.outer_daily,"prior"),["day"],["outer_spend"]),"day");
        const monthNum = Number(String(DATA?.meta?.current_start || "").slice(5,7)) || 4;
        return {
          labels: days.map(day => `${monthNum}/${day}`),
          cvrCurrent: days.map(day => { const row=currentShop.get(day)||{}; return ratio(row.orders,row.visitors); }),
          cvrPrior: days.map(day => { const row=priorShop.get(day)||{}; return ratio(row.orders,row.visitors); }),
          visitorsCurrent: days.map(day => safeNum((currentShop.get(day)||{}).visitors)),
          visitorsPrior: days.map(day => safeNum((priorShop.get(day)||{}).visitors)),
          outerCurrent: days.map(day => safeNum((currentOuter.get(day)||{}).outer_spend)),
          outerPrior: days.map(day => safeNum((priorOuter.get(day)||{}).outer_spend)),
          ordersCurrent: days.map(day => safeNum((currentShop.get(day)||{}).orders)),
          ordersPrior: days.map(day => safeNum((priorShop.get(day)||{}).orders))
        };
      }
      function shopTickLabelsV2(labels){
        return labels.filter((_,idx)=>idx===0||idx===labels.length-1||idx%2===1);
      }
      function ensureShopCardLayoutV2(){
        const leftCard = document.getElementById("shop-analysis-left-card");
        const rightCard = document.getElementById("shop-analysis-right-card");
        if(!leftCard || !rightCard) return;
        const shareBlock = document.getElementById("shop-share-chart")?.closest(".shop-block");
        const tableBlock = document.getElementById("shop-table-holder")?.closest(".shop-block");
        const cvrBlock = document.getElementById("shop-cvr-chart")?.closest(".shop-block");
        const spendBlock = document.getElementById("shop-spend-visitor-chart")?.closest(".shop-block");
        if(shareBlock?.querySelector(".shop-tag")) shareBlock.querySelector(".shop-tag").textContent = "店铺GMV占比";
        if(tableBlock?.querySelector(".shop-tag")) tableBlock.querySelector(".shop-tag").textContent = "店铺诊断表";
        if(cvrBlock?.querySelector(".shop-tag")) cvrBlock.querySelector(".shop-tag").textContent = "转化率日趋";
        if(spendBlock?.querySelector(".shop-tag")) spendBlock.querySelector(".shop-tag").textContent = "站外花费&流量";

        const stack = rightCard.querySelector(".shop-stack");
        const miniGrid = rightCard.querySelector(".shop-mini-grid");
        if(stack && !document.getElementById("shop-order-chart")){
          const orderBlock = document.createElement("div");
          orderBlock.className = "shop-block";
          orderBlock.innerHTML = '<div class="shop-block-head"><div class="shop-tag">订单日趋</div><button type="button" data-shop-download="shop-order-chart">下载图表</button></div><div id="shop-order-chart" class="plot shop-plot shop-plot-small"></div>';
          stack.insertBefore(orderBlock, miniGrid || null);
        }
        let topCard = document.getElementById("shop-top-products-card");
        if(!topCard){
          topCard = document.createElement("div");
          topCard.className = "card shop-analysis-card";
          topCard.id = "shop-top-products-card";
          topCard.innerHTML = '<div class="chart-head"><div><div class="section-label">Top Products</div><h2>店铺Top品</h2><div class="muted">查看 Authorize 和 Certified 的 GMV Top 产品</div></div><div class="shop-top-filter"><label for="shop-top-day-filter">单日筛选</label><select id="shop-top-day-filter"></select></div></div><div class="shop-top-card-body"></div>';
          rightCard.parentElement.insertBefore(topCard, rightCard.nextSibling);
        }
        if(miniGrid){
          const topBody = topCard.querySelector(".shop-top-card-body");
          if(topBody && miniGrid.parentElement !== topBody) topBody.appendChild(miniGrid);
        }
        const filter = document.getElementById("shop-top-day-filter");
        if(filter){
          if(state.shopTopDay===undefined) state.shopTopDay = "";
          const monthNum = Number(String(DATA?.meta?.current_start || "").slice(5,7)) || 4;
          const options = [{value:"",label:"全部日期"}].concat(Array.from({length:Number(DATA?.meta?.day_count||0)},(_,idx)=>({value:String(idx+1),label:`${monthNum}/${idx+1}`})));
          filter.innerHTML = options.map(opt=>`<option value="${opt.value}" ${String(state.shopTopDay)===String(opt.value)?"selected":""}>${opt.label}</option>`).join("");
          filter.onchange = function(e){
            state.shopTopDay = e.target.value;
            renderAll();
          };
        }
        document.querySelectorAll("[data-shop-download]").forEach(btn=>{
          btn.onclick=()=>downloadCanvas(btn.dataset.shopDownload);
        });
      }
      function buildShopTopRowsByDay(shopName){
        if(!shopName || typeof aggregateBy!=="function" || typeof filterRows!=="function") return [];
        const dayValue = String(state.shopTopDay || "");
        let rows = filterRows(DATA.shop_daily,"current").filter(row=>row.shop===shopName);
        if(dayValue) rows = rows.filter(row=>String(row.day)===dayValue);
        return aggregateBy(rows,["product"],["gmv","visitors","orders"]).map(row=>{
          const g=safeNum(row.gmv),v=safeNum(row.visitors),o=safeNum(row.orders);
          return {label:row.product||"未归类",gmv:g,visitors:v,orders:o,cvr:ratio(o,v)};
        }).filter(row=>row.label&&row.label!=="未归类").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,10);
      }
      function renderShopTopProductsByDay(){
        const topCard = document.getElementById("shop-top-products-card");
        if(!topCard) return;
        const muted = topCard.querySelector(".muted");
        if(muted) muted.textContent = `查看 Authorize 和 Certified 在 ${shopTopDayLabel(state.shopTopDay)} 的 GMV Top 产品`;
        [["shop-top-a","shop-top-a-title","Authorize"],["shop-top-b","shop-top-b-title","Certified"]].forEach(([holderId,titleId,shopName])=>{
          const holder=document.getElementById(holderId);
          const title=document.getElementById(titleId);
          if(!holder||!title) return;
          title.textContent = `${shopName}店Top10产品`;
          const rows = buildShopTopRowsByDay(shopName);
          if(!rows.length){
            holder.innerHTML = '<div class="shop-mini-empty">当前筛选下该店铺暂无有效产品数据。</div>';
            return;
          }
          holder.innerHTML = `<div class="shop-rank-list">${rows.map((row,idx)=>`
            <div class="shop-rank-item">
              <span class="shop-rank-no">${idx+1}</span>
              <div class="shop-rank-name">${row.label}</div>
              <div class="shop-rank-metric"><strong>${fmtMoney(row.gmv)}</strong><span>订单 ${fmtNum(row.orders)} | CVR ${fmtPct(row.cvr)}</span></div>
            </div>`).join("")}</div>`;
        });
      }
      function redrawShopOrderChartV2(){
        const holder = document.getElementById("shop-order-chart");
        if(!holder || typeof renderPlot!=="function") return;
        const daily = buildShopDailySeriesV2();
        if(!daily || !daily.labels.length){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无订单趋势数据。</div>';
          return;
        }
        const tickLabels = shopTickLabelsV2(daily.labels);
        const hasOrders = daily.ordersCurrent.concat(daily.ordersPrior).some(v=>safeNum(v)>0);
        if(!hasOrders){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无订单趋势数据。</div>';
          return;
        }
        renderPlot("shop-order-chart",[
          {type:"scatter",mode:"lines+markers",name:"4月订单",x:daily.labels,y:daily.ordersCurrent,line:{color:"#2f6bff",width:3},marker:{size:7,color:"#2f6bff"},hovertemplate:"<b>%{x}</b><br>4月订单: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"3月同期订单",x:daily.labels,y:daily.ordersPrior,line:{color:"#9ec6ff",width:3},marker:{size:7,color:"#9ec6ff"},hovertemplate:"<b>%{x}</b><br>3月同期订单: %{y:,.0f}<extra></extra>"}
        ],{
          height:240,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:60,r:18,t:38,b:72},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{title:"日期",tickmode:"array",tickvals:tickLabels,ticktext:tickLabels,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"订单",tickformat:",.0f",gridcolor:"#eef3fb",zeroline:false,tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }
      const prevRenderAllShopCardsV2 = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAllShopCardsV2 === "function") prevRenderAllShopCardsV2(reset);
        ensureShopCardLayoutV2();
        renderShopTopProductsByDay();
        redrawShopOrderChartV2();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(function(){ ensureShopCardLayoutV2(); renderShopTopProductsByDay(); redrawShopOrderChartV2(); }, 0); });
      }else{
        setTimeout(function(){ ensureShopCardLayoutV2(); renderShopTopProductsByDay(); redrawShopOrderChartV2(); }, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      function ensureShopShareToggleV2(){
        const leftCard = document.getElementById("shop-analysis-left-card");
        const tools = leftCard && leftCard.querySelector(".chart-tools");
        if(!leftCard || !tools) return;
        if(state.shopSharePeriod===undefined) state.shopSharePeriod = "current";
        let select = document.getElementById("shop-share-period");
        if(!select){
          select = document.createElement("select");
          select.id = "shop-share-period";
          select.style.height = "40px";
          select.style.borderRadius = "14px";
          select.style.border = "1px solid rgba(46,90,167,.16)";
          select.style.padding = "0 12px";
          select.style.background = "#fff";
          select.style.color = "#233456";
          select.style.font = "inherit";
          tools.insertBefore(select, tools.firstChild);
        }
        select.innerHTML = [
          `<option value="current" ${state.shopSharePeriod==="current"?"selected":""}>当前</option>`,
          `<option value="prior" ${state.shopSharePeriod==="prior"?"selected":""}>上月同期</option>`
        ].join("");
        select.onchange = function(e){
          state.shopSharePeriod = e.target.value;
          renderAll();
        };
      }
      function redrawShopShareByPeriodV2(){
        const holder = document.getElementById("shop-share-chart");
        if(!holder || typeof aggregateBy!=="function" || typeof filterRows!=="function" || typeof indexBy!=="function") return;
        const period = state.shopSharePeriod || "current";
        const rows = aggregateBy(filterRows(DATA.shop_daily, period),["shop"],["gmv"]).map(row=>({
          label: row.shop || "未归类",
          gmv: safeNum(row.gmv)
        })).filter(row=>row.label && row.label!=="未归类" && row.gmv>0).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
        if(!rows.length){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无店铺占比数据。</div>';
          return;
        }
        renderPlot("shop-share-chart",[
          {
            type:"pie",
            labels:rows.map(r=>r.label),
            values:rows.map(r=>safeNum(r.gmv)),
            hole:.58,
            sort:false,
            direction:"clockwise",
            textinfo:"percent",
            textposition:"inside",
            insidetextorientation:"radial",
            marker:{colors:rows.map((_,idx)=>["#2f6bff","#67a6ff","#9ec6ff","#7dd3fc","#dbeafe"][idx%5]),line:{width:0}},
            hovertemplate:"<b>%{label}</b><br>GMV: %{value:,.0f}<br>占比: %{percent}<extra></extra>"
          }
        ],{
          height:340,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:8,r:128,t:10,b:10},
          font:basePlotFont(),
          showlegend:true,
          legend:{orientation:"v",x:1.02,xanchor:"left",y:.5,yanchor:"middle",bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1,font:{family:basePlotFont().family,size:12,color:"#334155"}}
        },basePlotConfig());
      }
      const prevRenderAllShopShareV2 = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAllShopShareV2 === "function") prevRenderAllShopShareV2(reset);
        ensureShopShareToggleV2();
        redrawShopShareByPeriodV2();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(function(){ ensureShopShareToggleV2(); redrawShopShareByPeriodV2(); }, 0); });
      }else{
        setTimeout(function(){ ensureShopShareToggleV2(); redrawShopShareByPeriodV2(); }, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      const SHOP_COLORS_FINAL = ["#2f6bff","#67a6ff","#9ec6ff","#7dd3fc","#dbeafe"];
      function shopDayOptionsFinal(selectedValue, includeAll=false){
        const monthNum=Number(String(DATA?.meta?.current_start||"").slice(5,7))||4;
        const dayCount=Number(DATA?.meta?.day_count||0);
        const selectedText=selectedValue===undefined||selectedValue===null?"":String(selectedValue);
        const allOption=includeAll?`<option value="">全部日期</option>`:"";
        return `${allOption}${Array.from({length:dayCount},(_,idx)=>{const day=String(idx+1);return `<option value="${day}" ${selectedText===day?"selected":""}>${monthNum}/${day}</option>`;}).join("")}`;
      }
      function latestShopDayFinal(){
        const currentRows=filterRows(DATA.shop_daily,"current");
        const dayValues=currentRows.map(row=>Number(row.day)).filter(day=>Number.isFinite(day)&&day>0);
        if(dayValues.length) return String(Math.max(...dayValues));
        const fallback=Number(DATA?.meta?.day_count||0);
        return fallback>0?String(fallback):"";
      }
      function normalizeShopDaySelection(value, allowBlank=false){
        const text=value===undefined||value===null?"":String(value);
        if(!text) return allowBlank?"":latestShopDayFinal();
        const dayNum=Number(text);
        const maxDay=Number(DATA?.meta?.day_count||0);
        if(!Number.isFinite(dayNum)||dayNum<1||dayNum>maxDay) return allowBlank?"":latestShopDayFinal();
        return String(dayNum);
      }
      function buildShopRowsFinal(dayValue){
        const currentRows=filterRows(DATA.shop_daily,"current").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const priorRows=filterRows(DATA.shop_daily,"prior").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const currentOuterRows=filterRows(DATA.outer_daily,"current").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const priorOuterRows=filterRows(DATA.outer_daily,"prior").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const currentInnerRows=filterRows(DATA.inner_daily,"current").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const priorInnerRows=filterRows(DATA.inner_daily,"prior").filter(row=>!dayValue||String(row.day)===String(dayValue));
        const cs=aggregateBy(currentRows,["shop"],["gmv","visitors","orders"]);
        const ps=aggregateBy(priorRows,["shop"],["gmv","visitors","orders"]);
        const co=aggregateBy(currentOuterRows,["shop"],["outer_spend"]);
        const po=aggregateBy(priorOuterRows,["shop"],["outer_spend"]);
        const ci=aggregateBy(currentInnerRows,["shop"],["inner_spend"]);
        const pi=aggregateBy(priorInnerRows,["shop"],["inner_spend"]);
        const csMap=indexBy(cs,"shop"),psMap=indexBy(ps,"shop"),coMap=indexBy(co,"shop"),poMap=indexBy(po,"shop"),ciMap=indexBy(ci,"shop"),piMap=indexBy(pi,"shop");
        const keys=new Set([...csMap.keys(),...psMap.keys(),...coMap.keys(),...poMap.keys(),...ciMap.keys(),...piMap.keys()]);
        return Array.from(keys).map(key=>{
          const cur={...(csMap.get(key)||{}),...(coMap.get(key)||{}),...(ciMap.get(key)||{})};
          const pre={...(psMap.get(key)||{}),...(poMap.get(key)||{}),...(piMap.get(key)||{})};
          const g=safeNum(cur.gmv), pg=safeNum(pre.gmv);
          const v=safeNum(cur.visitors), pv=safeNum(pre.visitors);
          const o=safeNum(cur.orders), poOrders=safeNum(pre.orders);
          const inner=safeNum(cur.inner_spend), pInner=safeNum(pre.inner_spend);
          const outer=safeNum(cur.outer_spend), pOuter=safeNum(pre.outer_spend);
          return {
            label:key,
            gmv:g, priorGmv:pg, gmvMom:pctChange(g,pg),
            visitors:v, priorVisitors:pv, visitorsMom:pctChange(v,pv),
            orders:o, priorOrders:poOrders, ordersMom:pctChange(o,poOrders),
            cvr:ratio(o,v), priorCvr:ratio(poOrders,pv), cvrMom:pctChange(ratio(o,v),ratio(poOrders,pv)),
            innerSpend:inner, priorInnerSpend:pInner, innerSpendMom:pctChange(inner,pInner),
            outerSpend:outer, priorOuterSpend:pOuter, outerSpendMom:pctChange(outer,pOuter)
          };
        }).filter(row=>row.label&&row.label!=="未归类"&&row.label!=="其他").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      }
      function buildShopDailySeriesFinalV3(){
        const dayCount=Number(DATA?.meta?.day_count||0);
        const days=Array.from({length:dayCount},(_,i)=>i+1);
        const currentShop=indexBy(aggregateBy(filterRows(DATA.shop_daily,"current"),["day"],["visitors","orders"]),"day");
        const priorShop=indexBy(aggregateBy(filterRows(DATA.shop_daily,"prior"),["day"],["visitors","orders"]),"day");
        const currentOuter=indexBy(aggregateBy(filterRows(DATA.outer_daily,"current"),["day"],["outer_spend"]),"day");
        const priorOuter=indexBy(aggregateBy(filterRows(DATA.outer_daily,"prior"),["day"],["outer_spend"]),"day");
        const currentInner=indexBy(aggregateBy(filterRows(DATA.inner_daily,"current"),["day"],["inner_spend"]),"day");
        const priorInner=indexBy(aggregateBy(filterRows(DATA.inner_daily,"prior"),["day"],["inner_spend"]),"day");
        const monthNum=Number(String(DATA?.meta?.current_start||"").slice(5,7))||4;
        return {
          labels:days.map(day=>`${monthNum}/${day}`),
          cvrCurrent:days.map(day=>{const row=currentShop.get(day)||{};return ratio(row.orders,row.visitors);}),
          cvrPrior:days.map(day=>{const row=priorShop.get(day)||{};return ratio(row.orders,row.visitors);}),
          visitorsCurrent:days.map(day=>safeNum((currentShop.get(day)||{}).visitors)),
          visitorsPrior:days.map(day=>safeNum((priorShop.get(day)||{}).visitors)),
          ordersCurrent:days.map(day=>safeNum((currentShop.get(day)||{}).orders)),
          ordersPrior:days.map(day=>safeNum((priorShop.get(day)||{}).orders)),
          innerCurrent:days.map(day=>safeNum((currentInner.get(day)||{}).inner_spend)),
          innerPrior:days.map(day=>safeNum((priorInner.get(day)||{}).inner_spend)),
          outerCurrent:days.map(day=>safeNum((currentOuter.get(day)||{}).outer_spend)),
          outerPrior:days.map(day=>safeNum((priorOuter.get(day)||{}).outer_spend))
        };
      }
      function buildPieRowsFinal(dayValue){
        const rows=filterRows(DATA.shop_daily,"current").filter(row=>!dayValue||String(row.day)===String(dayValue));
        return aggregateBy(rows,["shop"],["gmv"]).map(row=>({
          label:row.shop||"未归类",
          gmv:safeNum(row.gmv)
        })).filter(row=>row.label&&row.label!=="未归类"&&row.gmv>0).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      }
      function buildShopTopRowsFinal(shopName){
        const dayValue=String(state.shopTopDay||"");
        let rows=filterRows(DATA.shop_daily,"current").filter(row=>row.shop===shopName);
        if(dayValue) rows=rows.filter(row=>String(row.day)===dayValue);
        return aggregateBy(rows,["product"],["gmv","visitors","orders"]).map(row=>{
          const g=safeNum(row.gmv),v=safeNum(row.visitors),o=safeNum(row.orders);
          return {label:row.product||"未归类",gmv:g,visitors:v,orders:o,cvr:ratio(o,v)};
        }).filter(row=>row.label&&row.label!=="未归类").sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv)).slice(0,10);
      }
      function shopTickLabelsFinal(labels){
        return labels.filter((_,idx)=>idx===0||idx===labels.length-1||idx%2===1);
      }
      function shopMomClassFinal(value){
        const n=safeNum(value);
        if(n>0.005) return "is-up";
        if(n<-0.005) return "is-down";
        return "is-flat";
      }
      function shopMomHtmlFinal(value){
        return `<span class="shop-mom ${shopMomClassFinal(value)}">${fmtDelta(value)}</span>`;
      }
      function monthDayLabelFinal(dayValue){
        const monthNum=Number(String(DATA?.meta?.current_start||"").slice(5,7))||4;
        return dayValue?`${monthNum}/${dayValue}`:"全部日期";
      }
      function mountShopCardsFinal(){
        const spendCard=document.getElementById("spend-mix")?.closest(".card");
        const grid=spendCard&&spendCard.parentElement;
        if(!grid||!spendCard) return null;
        ["shop-analysis-left-card","shop-analysis-right-card","shop-top-products-card"].forEach(id=>{
          document.getElementById(id)?.remove();
        });
        const left=document.createElement("div");
        left.className="card shop-analysis-card";
        left.id="shop-analysis-left-card";
        left.innerHTML='<div class="chart-head"><div><div class="section-label">Shop Analysis</div><h2>店铺分析</h2><div class="muted">聚焦各店铺贡献结构与核心经营诊断</div></div><div class="chart-tools"><div class="shop-top-filter"><label for="shop-share-day-filter">单日筛选</label><select id="shop-share-day-filter"></select></div><button type="button" data-shop-download="shop-share-chart">下载图表</button></div></div><div class="shop-stack"><div class="shop-block"><div class="shop-tag">店铺GMV占比</div><div id="shop-share-chart" class="plot shop-plot shop-plot-medium"></div></div><div class="shop-block"><div class="shop-block-head"><div class="shop-tag">店铺诊断表</div><div class="shop-top-filter"><label for="shop-table-day-filter">单日筛选</label><select id="shop-table-day-filter"></select></div></div><div id="shop-table-holder" class="shop-table-wrap"></div></div><div class="shop-block shop-cvr-lower-block"><div class="shop-block-head"><div class="shop-tag">转化率日趋</div><button type="button" data-shop-download="shop-cvr-chart">下载图表</button></div><div id="shop-cvr-chart" class="plot shop-plot shop-plot-small"></div></div></div><div class="narrative" id="shop-left-narrative"></div>';
        const right=document.createElement("div");
        right.className="card shop-analysis-card";
        right.id="shop-analysis-right-card";
        right.innerHTML='<div class="chart-head"><div><div class="section-label">Shop Conversion</div><h2>店铺承接分析</h2><div class="muted">看店铺转化、花费承接和主力产品结构</div></div></div><div class="shop-stack"><div class="shop-block"><div class="shop-block-head"><div class="shop-tag">站内花费&访客</div><button type="button" data-shop-download="shop-inner-visitor-chart">下载图表</button></div><div id="shop-inner-visitor-chart" class="plot shop-plot shop-plot-medium"></div></div><div class="shop-block"><div class="shop-block-head"><div class="shop-tag">站外花费&流量</div><button type="button" data-shop-download="shop-spend-visitor-chart">下载图表</button></div><div id="shop-spend-visitor-chart" class="plot shop-plot shop-plot-medium"></div></div><div class="shop-block"><div class="shop-block-head"><div class="shop-tag">订单日趋</div><button type="button" data-shop-download="shop-order-chart">下载图表</button></div><div id="shop-order-chart" class="plot shop-plot shop-plot-small"></div></div></div><div class="narrative" id="shop-right-narrative"></div>';
        const top=document.createElement("div");
        top.className="card shop-analysis-card";
        top.id="shop-top-products-card";
        top.style.gridColumn="2";
        top.innerHTML='<div class="chart-head"><div><div class="section-label">Top Products</div><h2>店铺Top品</h2><div class="muted" id="shop-top-muted">查看 Authorize 和 Certified 的 GMV Top 产品</div></div><div class="shop-top-filter"><label for="shop-top-day-filter">单日筛选</label><select id="shop-top-day-filter"></select></div></div><div class="shop-mini-grid"><div class="shop-rank-wrap"><div class="shop-tag" id="shop-top-a-title">Authorize店Top10产品</div><div id="shop-top-a"></div></div><div class="shop-rank-wrap"><div class="shop-tag" id="shop-top-b-title">Certified店Top10产品</div><div id="shop-top-b"></div></div></div>';
        grid.insertBefore(left, spendCard.nextSibling);
        grid.insertBefore(right, left.nextSibling);
        grid.insertBefore(top, right.nextSibling);
        document.querySelectorAll("[data-shop-download]").forEach(btn=>{
          btn.onclick=()=>downloadCanvas(btn.dataset.shopDownload);
        });
        return {left,right,top};
      }
      function renderShopTableFinal(rows){
        const holder=document.getElementById("shop-table-holder");
        if(!holder) return;
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺诊断表数据。</div>';
          return;
        }
        holder.innerHTML=`<table><thead><tr><th>店铺</th><th>GMV</th><th>环比</th><th>访客</th><th>环比</th><th>转化率</th><th>环比</th><th>站内花费</th><th>环比</th><th>站外花费</th><th>环比</th></tr></thead><tbody>${rows.map(row=>`<tr><td>${row.label}</td><td>${fmtMoney(row.gmv)}</td><td>${shopMomHtmlFinal(row.gmvMom)}</td><td>${fmtNum(row.visitors)}</td><td>${shopMomHtmlFinal(row.visitorsMom)}</td><td>${fmtPct(row.cvr)}</td><td>${shopMomHtmlFinal(row.cvrMom)}</td><td>${fmtMoney(row.innerSpend)}</td><td>${shopMomHtmlFinal(row.innerSpendMom)}</td><td>${fmtMoney(row.outerSpend)}</td><td>${shopMomHtmlFinal(row.outerSpendMom)}</td></tr>`).join("")}</tbody></table>`;
      }
      function renderShopShareFinal(){
        const select=document.getElementById("shop-share-day-filter");
        const holder=document.getElementById("shop-share-chart");
        if(!select || !holder) return;
        state.shopShareDay=normalizeShopDaySelection(state.shopShareDay,true);
        select.innerHTML=shopDayOptionsFinal(state.shopShareDay,true);
        select.onchange=function(e){ state.shopShareDay=normalizeShopDaySelection(e.target.value,true); renderShopCardsFinal(); };
        const rows=buildPieRowsFinal(state.shopShareDay);
        if(!rows.length){
          holder.innerHTML='<div class="empty-state">当前筛选下暂无店铺占比数据。</div>';
          return;
        }
        renderPlot("shop-share-chart",[{type:"pie",labels:rows.map(r=>r.label),values:rows.map(r=>safeNum(r.gmv)),hole:.58,sort:false,direction:"clockwise",textinfo:"percent",textposition:"inside",insidetextorientation:"radial",marker:{colors:rows.map((_,idx)=>SHOP_COLORS_FINAL[idx%SHOP_COLORS_FINAL.length]),line:{width:0}},hovertemplate:"<b>%{label}</b><br>GMV: %{value:,.0f}<br>占比: %{percent}<extra></extra>"}],{height:340,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:8,r:128,t:10,b:10},font:basePlotFont(),showlegend:true,legend:{orientation:"v",x:1.02,xanchor:"left",y:.5,yanchor:"middle",bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1,font:{family:basePlotFont().family,size:12,color:"#334155"}}},basePlotConfig());
      }
      function renderShopTrendChartsFinal(){
        const daily=buildShopDailySeriesFinalV3();
        if(!daily||!daily.labels.length) return;
        const ticks=shopTickLabelsFinal(daily.labels);
        renderPlot("shop-inner-visitor-chart",[
          {type:"bar",name:"4月站内花费",x:daily.labels,y:daily.innerCurrent,marker:{color:"#2f6bff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>4月站内花费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"3月同期站内花费",x:daily.labels,y:daily.innerPrior,marker:{color:"#b7d5ff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>3月同期站内花费: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"4月访客",x:daily.labels,y:daily.visitorsCurrent,yaxis:"y2",line:{color:"#f59e0b",width:3},marker:{size:7,color:"#f59e0b"},hovertemplate:"<b>%{x}</b><br>4月访客: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"3月同期访客",x:daily.labels,y:daily.visitorsPrior,yaxis:"y2",line:{color:"#fcd34d",width:3},marker:{size:7,color:"#fcd34d"},hovertemplate:"<b>%{x}</b><br>3月同期访客: %{y:,.0f}<extra></extra>"}
        ],{height:340,barmode:"group",bargap:.22,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:64,r:62,t:46,b:72},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",x:0,y:1.2,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"array",tickvals:ticks,ticktext:ticks,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"站内花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:"访客",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());
        renderPlot("shop-cvr-chart",[
          {type:"scatter",mode:"lines+markers",name:"4月CVR",x:daily.labels,y:daily.cvrCurrent,line:{color:"#2f6bff",width:3},marker:{size:7,color:"#2f6bff"},hovertemplate:"<b>%{x}</b><br>4月CVR: %{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"3月同期CVR",x:daily.labels,y:daily.cvrPrior,line:{color:"#9ec6ff",width:3},marker:{size:7,color:"#9ec6ff"},hovertemplate:"<b>%{x}</b><br>3月同期CVR: %{y:.1%}<extra></extra>"}
        ],{height:240,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:60,r:18,t:38,b:72},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"array",tickvals:ticks,ticktext:ticks,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"CVR",tickformat:".0%",gridcolor:"#eef3fb",zeroline:false,tickfont:{color:"#64748b"}}},basePlotConfig());
        renderPlot("shop-spend-visitor-chart",[
          {type:"bar",name:"4月站外花费",x:daily.labels,y:daily.outerCurrent,marker:{color:"#2f6bff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>4月站外花费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"3月同期站外花费",x:daily.labels,y:daily.outerPrior,marker:{color:"#b7d5ff",line:{width:0}},hovertemplate:"<b>%{x}</b><br>3月同期站外花费: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"4月访客",x:daily.labels,y:daily.visitorsCurrent,yaxis:"y2",line:{color:"#f59e0b",width:3},marker:{size:7,color:"#f59e0b"},hovertemplate:"<b>%{x}</b><br>4月访客: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"3月同期访客",x:daily.labels,y:daily.visitorsPrior,yaxis:"y2",line:{color:"#fcd34d",width:3},marker:{size:7,color:"#fcd34d"},hovertemplate:"<b>%{x}</b><br>3月同期访客: %{y:,.0f}<extra></extra>"}
        ],{height:340,barmode:"group",bargap:.22,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:64,r:62,t:46,b:72},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",x:0,y:1.2,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"array",tickvals:ticks,ticktext:ticks,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"站外花费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},yaxis2:{title:"访客",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}},basePlotConfig());
        renderPlot("shop-order-chart",[
          {type:"scatter",mode:"lines+markers",name:"4月订单",x:daily.labels,y:daily.ordersCurrent,line:{color:"#2f6bff",width:3},marker:{size:7,color:"#2f6bff"},hovertemplate:"<b>%{x}</b><br>4月订单: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"3月同期订单",x:daily.labels,y:daily.ordersPrior,line:{color:"#9ec6ff",width:3},marker:{size:7,color:"#9ec6ff"},hovertemplate:"<b>%{x}</b><br>3月同期订单: %{y:,.0f}<extra></extra>"}
        ],{height:240,paper_bgcolor:"rgba(0,0,0,0)",plot_bgcolor:"#ffffff",margin:{l:60,r:18,t:38,b:72},hovermode:"x unified",hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},font:basePlotFont(),legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},xaxis:{title:"日期",tickmode:"array",tickvals:ticks,ticktext:ticks,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},yaxis:{title:"订单",tickformat:",.0f",gridcolor:"#eef3fb",zeroline:false,tickfont:{color:"#64748b"}}},basePlotConfig());
      }
      function renderShopTopFinal(){
        const select=document.getElementById("shop-top-day-filter");
        const muted=document.getElementById("shop-top-muted");
        if(!select) return;
        if(state.shopTopDay===undefined) state.shopTopDay="";
        const monthNum=Number(String(DATA?.meta?.current_start||"").slice(5,7))||4;
        select.innerHTML=`<option value="">全部日期</option>${Array.from({length:Number(DATA?.meta?.day_count||0)},(_,idx)=>`<option value="${idx+1}" ${String(state.shopTopDay)===String(idx+1)?"selected":""}>${monthNum}/${idx+1}</option>`).join("")}`;
        select.onchange=function(e){ state.shopTopDay=e.target.value; renderShopCardsFinal(); };
        if(muted) muted.textContent=`查看 Authorize 和 Certified 在 ${monthDayLabelFinal(state.shopTopDay)} 的 GMV Top 产品`;
        [["shop-top-a","Authorize"],["shop-top-b","Certified"]].forEach(([holderId,shopName])=>{
          const holder=document.getElementById(holderId);
          if(!holder) return;
          const rows=buildShopTopRowsFinal(shopName);
          if(!rows.length){
            holder.innerHTML='<div class="shop-mini-empty">当前筛选下该店铺暂无有效产品数据。</div>';
            return;
          }
          holder.innerHTML=`<div class="shop-rank-list">${rows.map((row,idx)=>`<div class="shop-rank-item"><span class="shop-rank-no">${idx+1}</span><div class="shop-rank-name">${row.label}</div><div class="shop-rank-metric"><strong>${fmtMoney(row.gmv)}</strong><span>订单 ${fmtNum(row.orders)} | CVR ${fmtPct(row.cvr)}</span></div></div>`).join("")}</div>`;
        });
      }
      function renderShopNarrativesFinal(rows){
        const left=document.getElementById("shop-left-narrative");
        const right=document.getElementById("shop-right-narrative");
        if(!left||!right) return;
        if(!rows.length){
          left.innerHTML='<div class="empty-state">当前筛选下暂无店铺分析结论。</div>';
          right.innerHTML='<div class="empty-state">当前筛选下暂无店铺承接分析结论。</div>';
          return;
        }
        const totalGmv=rows.reduce((sum,row)=>sum+safeNum(row.gmv),0);
        const top=rows[0];
        const bestCvr=[...rows].sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const spendLead=[...rows].sort((a,b)=>safeNum(b.outerSpend)-safeNum(a.outerSpend))[0];
        left.innerHTML=`<p>${top.label} 当前店铺GMV ${fmtMoney(top.gmv)}，贡献占比 ${fmtPct(ratio(top.gmv,totalGmv))}。</p><p>${bestCvr.label} 的CVR最高，为 ${fmtPct(bestCvr.cvr)}，适合作为高承接店铺参考。</p>`;
        right.innerHTML=`<p>${spendLead.label} 当前站外花费最高，为 ${fmtMoney(spendLead.outerSpend)}，对应访客 ${fmtNum(spendLead.visitors)}，订单 ${fmtNum(spendLead.orders)}。</p><p><strong>下一步：</strong>优先看高花费店铺的人群匹配和详情页承接，再决定是否继续放量。</p>`;
      }
      function renderShopCardsFinal(){
        mountShopCardsFinal();
        const tableSelect=document.getElementById("shop-table-day-filter");
        state.shopTableDay=normalizeShopDaySelection(state.shopTableDay,true);
        if(tableSelect){
          tableSelect.innerHTML=shopDayOptionsFinal(state.shopTableDay,true);
          tableSelect.onchange=function(e){ state.shopTableDay=normalizeShopDaySelection(e.target.value,true); renderShopCardsFinal(); };
        }
        const rows=buildShopRowsFinal(state.shopTableDay);
        renderShopShareFinal();
        renderShopTableFinal(rows);
        renderShopTrendChartsFinal();
        renderShopTopFinal();
        renderShopNarrativesFinal(rows);
      }
      const prevRenderAllShopCardsFinal = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAllShopCardsFinal === "function") prevRenderAllShopCardsFinal(reset);
        renderShopCardsFinal();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(renderShopCardsFinal, 0); });
      }else{
        setTimeout(renderShopCardsFinal, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    #bundle-chart.cardified{}
    #bundle-chart-holder,
    .bundle-left-stack{
      grid-area:plot;
      min-width:0;
    }
    .bundle-left-stack{
      display:grid;
      grid-template-rows:auto auto;
      gap:18px;
      align-self:start;
    }
    .bundle-table-card{
      border:1px solid rgba(46,90,167,.12);
      border-radius:20px;
      background:#f8fbff;
      padding:18px 18px 12px;
      overflow:hidden;
    }
    .bundle-table-head{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      margin-bottom:12px;
    }
    .bundle-table-title{
      display:inline-flex;
      align-items:center;
      min-height:38px;
      padding:8px 16px;
      border-radius:12px;
      background:#4d98ff;
      color:#fff;
      font-size:14px;
      font-weight:700;
      line-height:1.5;
    }
    .bundle-table-note{
      color:#6b7f9f;
      font-size:12px;
      white-space:nowrap;
    }
    .bundle-table-wrap{
      overflow-x:auto;
      overflow-y:hidden;
      border-radius:16px;
      border:1px solid rgba(148,163,184,.14);
      background:#fff;
    }
    .bundle-table{
      width:100%;
      min-width:1280px;
      border-collapse:separate;
      border-spacing:0;
    }
    .bundle-table th,
    .bundle-table td{
      padding:12px 14px;
      border-bottom:1px solid #edf2f9;
      text-align:left;
      font-size:13px;
      white-space:nowrap;
      vertical-align:middle;
    }
    .bundle-table th{
      position:sticky;
      top:0;
      background:#f5f9ff;
      color:#31507e;
      font-weight:700;
      z-index:1;
    }
    .bundle-table tr:last-child td{
      border-bottom:none;
    }
    .bundle-table .mom{
      font-weight:800;
    }
    .bundle-table .mom.pos{
      color:#ef4444;
    }
    .bundle-table .mom.neg{
      color:#16a34a;
    }
    .bundle-table .mom.flat{
      color:#64748b;
    }
    .bundle-empty{
      padding:18px 12px;
      color:#7b8ba5;
      font-size:13px;
      line-height:1.7;
    }
    @media (max-width:1200px){
      .bundle-left-stack{
        grid-template-rows:auto auto;
      }
      .bundle-table{
        min-width:1080px;
      }
    }
  </style>
  <script>
    (function(){
      function bundleMomClass(value){
        const n = safeNum(value);
        if(n > 0.005) return "pos";
        if(n < -0.005) return "neg";
        return "flat";
      }
      function ensureBundleLayoutFinal(){
        const bundleCard = document.getElementById("bundle-chart")?.closest(".card");
        const plot = document.getElementById("bundle-chart");
        if(!bundleCard || !plot) return null;
        let stack = document.getElementById("bundle-left-stack");
        if(!stack){
          stack = document.createElement("div");
          stack.id = "bundle-left-stack";
          stack.className = "bundle-left-stack";
          plot.parentElement.insertBefore(stack, plot);
          stack.appendChild(plot);
        }
        let tableCard = document.getElementById("bundle-table-card");
        if(!tableCard){
          tableCard = document.createElement("div");
          tableCard.id = "bundle-table-card";
          tableCard.className = "bundle-table-card";
          tableCard.innerHTML = '<div class="bundle-table-head"><div class="bundle-table-title">新套组诊断表</div><div class="bundle-table-note">字段过多时支持左右滑动</div></div><div class="bundle-table-wrap" id="bundle-table-wrap"></div>';
          stack.appendChild(tableCard);
        }
        return tableCard;
      }
      function renderBundleTableFinal(){
        const tableCard = ensureBundleLayoutFinal();
        const holder = document.getElementById("bundle-table-wrap");
        if(!tableCard || !holder || typeof buildBundleRows !== "function") return;
        const rows = buildBundleRows();
        if(!rows || !rows.length){
          holder.innerHTML = '<div class="bundle-empty">当前筛选下暂无重点新套组诊断数据。</div>';
          return;
        }
        holder.innerHTML = `<table class="bundle-table"><thead><tr><th>套组</th><th>GMV</th><th>环比</th><th>访客</th><th>环比</th><th>转化率</th><th>环比</th><th>站内花费</th><th>环比</th><th>站外花费</th><th>环比</th><th>站外曝光量</th><th>环比</th></tr></thead><tbody>${rows.map(row=>{
          const gmvMom = safeNum(row.gmvMom);
          const visitorsMom = safeNum(row.visitorsMom);
          const cvrMom = safeNum(row.cvrMom);
          const innerMom = pctChange(safeNum(row.innerSpend), safeNum(row.priorInnerSpend));
          const outerMom = pctChange(safeNum(row.outerSpend), safeNum(row.priorOuterSpend));
          const impMom = pctChange(safeNum(row.impressions), safeNum(row.priorImpressions));
          return `<tr>
            <td>${row.label}</td>
            <td>${fmtMoney(row.gmv)}</td>
            <td><span class="mom ${bundleMomClass(gmvMom)}">${fmtDelta(gmvMom)}</span></td>
            <td>${fmtNum(row.visitors)}</td>
            <td><span class="mom ${bundleMomClass(visitorsMom)}">${fmtDelta(visitorsMom)}</span></td>
            <td>${fmtPct(row.cvr)}</td>
            <td><span class="mom ${bundleMomClass(cvrMom)}">${fmtDelta(cvrMom)}</span></td>
            <td>${fmtMoney(row.innerSpend)}</td>
            <td><span class="mom ${bundleMomClass(innerMom)}">${fmtDelta(innerMom)}</span></td>
            <td>${fmtMoney(row.outerSpend)}</td>
            <td><span class="mom ${bundleMomClass(outerMom)}">${fmtDelta(outerMom)}</span></td>
            <td>${fmtNum(row.impressions)}</td>
            <td><span class="mom ${bundleMomClass(impMom)}">${fmtDelta(impMom)}</span></td>
          </tr>`;
        }).join("")}</tbody></table>`;
      }
      const previousRenderAllBundleTableFinal = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllBundleTableFinal === "function") previousRenderAllBundleTableFinal(reset);
        renderBundleTableFinal();
      };
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(renderBundleTableFinal, 0); });
      }else{
        setTimeout(renderBundleTableFinal, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .chart-grid{
      align-items:stretch;
    }
    .chart-grid > .card.shop-analysis-card{
      display:flex;
      flex-direction:column;
      height:100%;
      align-self:stretch;
    }
    #shop-analysis-left-card .shop-stack,
    #shop-analysis-right-card .shop-stack{
      display:flex;
      flex-direction:column;
      gap:18px;
      flex:1 1 auto;
    }
    #shop-analysis-left-card .shop-block,
    #shop-analysis-right-card .shop-block{
      display:flex;
      flex-direction:column;
    }
    #shop-analysis-left-card .shop-cvr-lower-block{
      margin-top:22px;
    }
    #shop-analysis-left-card .narrative,
    #shop-analysis-right-card .narrative{
      margin-top:18px;
      min-height:152px;
      display:flex;
      flex-direction:column;
      justify-content:flex-start;
    }
    .chart-grid > .card.chart-grid-force-wide{
      grid-column:1 / -1 !important;
      display:grid !important;
      grid-template-columns:minmax(0,1.24fr) minmax(320px,.76fr) !important;
      grid-template-areas:"head head" "plot narrative" !important;
      gap:18px 22px !important;
      align-items:start !important;
    }
    .chart-grid > .card.chart-grid-force-wide > .chart-head{
      grid-area:head;
      margin-bottom:0;
    }
    .chart-grid > .card.chart-grid-force-wide > .plot,
    .chart-grid > .card.chart-grid-force-wide > canvas{
      grid-area:plot;
      min-width:0;
      margin-top:0;
    }
    .chart-grid > .card.chart-grid-force-wide > .narrative{
      grid-area:narrative;
      margin-top:0;
      align-self:stretch;
      height:100%;
      min-height:100%;
      padding:22px 24px;
    }
    #shop-top-products-card.chart-grid-force-wide > .shop-mini-grid{
      grid-area:plot;
      min-width:0;
      display:grid;
      grid-template-columns:repeat(2,minmax(0,1fr));
      gap:18px;
      margin-top:0;
    }
    #shop-top-products-card.chart-grid-force-wide > .narrative{
      background:#f6fbff;
      border:1px solid var(--line);
      border-radius:20px;
    }
    #shop-top-products-card.chart-grid-force-wide .shop-rank-wrap{
      height:100%;
    }
    @media (max-width:1200px){
      .chart-grid > .card.chart-grid-force-wide{
        grid-template-columns:1fr !important;
        grid-template-areas:"head" "plot" "narrative" !important;
      }
      #shop-top-products-card.chart-grid-force-wide > .shop-mini-grid{
        grid-template-columns:1fr;
      }
      .chart-grid > .card.chart-grid-force-wide > .narrative{
        min-height:0;
        height:auto;
      }
    }
  </style>
  <script>
    (function(){
      function ensureShopTopNarrativeMount(){
        const topCard = document.getElementById("shop-top-products-card");
        if(!topCard) return null;
        let narrative = document.getElementById("shop-top-narrative");
        if(!narrative){
          narrative = document.createElement("div");
          narrative.className = "narrative";
          narrative.id = "shop-top-narrative";
          topCard.appendChild(narrative);
        }
        return narrative;
      }
      function parseShopTopBlock(holderId, shopName){
        const holder = document.getElementById(holderId);
        if(!holder) return null;
        const items = Array.from(holder.querySelectorAll(".shop-rank-item"));
        if(!items.length) return null;
        const topItem = items[0];
        const name = topItem.querySelector(".shop-rank-name")?.textContent?.trim() || "";
        const gmv = topItem.querySelector(".shop-rank-metric strong")?.textContent?.trim() || "";
        const metric = topItem.querySelector(".shop-rank-metric span")?.textContent?.trim() || "";
        return {shopName, name, gmv, metric, count:items.length};
      }
      function renderShopTopNarrativeFinal(){
        const holder = ensureShopTopNarrativeMount();
        if(!holder) return;
        const auth = parseShopTopBlock("shop-top-a","Authorize");
        const cert = parseShopTopBlock("shop-top-b","Certified");
        if(!auth && !cert){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无店铺 Top 品分析。</div>';
          return;
        }
        const paragraphs = [];
        if(auth){
          paragraphs.push(`${auth.name} 当前排在 Authorize 第一位，GMV 为 ${auth.gmv}，${auth.metric}。`);
        }
        if(cert){
          paragraphs.push(`${cert.name} 当前排在 Certified 第一位，GMV 为 ${cert.gmv}，${cert.metric}。`);
        }
        const todo = [];
        if(auth && cert){
          todo.push("继续结合单日筛选追踪两店 Top1 和 Top3 的位次变化，识别是投放切换还是自然动销在拉动结构变化。");
        }
        if(auth){
          todo.push("Authorize 侧优先盯头部品后的第二梯队承接品，避免流量继续过度堆在单一主力品。");
        }
        if(cert){
          todo.push("Certified 侧继续围绕高转化头部品补素材和联投组合，把榜单优势放大到更高 GMV 体量。");
        }
        holder.innerHTML = `<p>${paragraphs.join(" ")}</p><p><strong>下一步：</strong></p>${narrativeList(todo)}`;
      }
      function syncShopNarrativeHeights(){
        const left = document.getElementById("shop-left-narrative");
        const right = document.getElementById("shop-right-narrative");
        if(!left || !right) return;
        left.style.minHeight = "";
        right.style.minHeight = "";
        const target = Math.max(left.offsetHeight || 0, right.offsetHeight || 0, 152);
        left.style.minHeight = `${target}px`;
        right.style.minHeight = `${target}px`;
      }
      function applyFinalWideCardLayout(){
        const topCard = document.getElementById("shop-top-products-card");
        const ttGapCard = document.getElementById("tt-gap-chart")?.closest(".card");
        const bundleCard = document.getElementById("bundle-chart")?.closest(".card");
        [topCard, ttGapCard, bundleCard].forEach(card=>{
          if(card){
            card.classList.add("chart-grid-single-fill","chart-grid-force-wide");
          }
        });
        if(window.Plotly && window.Plotly.Plots && typeof window.Plotly.Plots.resize === "function"){
          [ttGapCard, bundleCard].forEach(card=>{
            const plot = card && card.querySelector(".plot");
            if(plot){
              try{ window.Plotly.Plots.resize(plot); }catch(err){}
            }
          });
        }
      }
      function applyShopWideCardFinal(){
        ensureShopTopNarrativeMount();
        renderShopTopNarrativeFinal();
        applyFinalWideCardLayout();
        setTimeout(syncShopNarrativeHeights, 0);
      }
      const previousRenderAllShopWideCardFinal = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllShopWideCardFinal === "function") previousRenderAllShopWideCardFinal(reset);
        applyShopWideCardFinal();
      };
      let finalWideCardTimer = null;
      window.addEventListener("resize", function(){
        clearTimeout(finalWideCardTimer);
        finalWideCardTimer = setTimeout(applyShopWideCardFinal, 120);
      });
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(applyShopWideCardFinal, 0); });
      }else{
        setTimeout(applyShopWideCardFinal, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .chart-grid > .card.chart-grid-force-wide > .bundle-left-stack{
      grid-area:plot;
      min-width:0;
    }
    .bundle-left-stack{
      display:grid;
      grid-template-rows:auto auto;
      gap:18px;
      align-self:start;
    }
    .bundle-table-card{
      border:1px solid rgba(46,90,167,.12);
      border-radius:20px;
      background:#f8fbff;
      padding:18px 18px 12px;
      overflow:hidden;
    }
    .bundle-table-head{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      margin-bottom:12px;
    }
    .bundle-table-title{
      display:inline-flex;
      align-items:center;
      min-height:38px;
      padding:8px 16px;
      border-radius:12px;
      background:#4d98ff;
      color:#ffffff;
      font-size:14px;
      font-weight:700;
      line-height:1.5;
    }
    .bundle-table-note{
      color:#6b7f9f;
      font-size:12px;
      white-space:nowrap;
    }
    .bundle-table-wrap{
      overflow-x:auto;
      overflow-y:hidden;
      border-radius:16px;
      border:1px solid rgba(148,163,184,.14);
      background:#ffffff;
    }
    .bundle-table{
      width:100%;
      min-width:1280px;
      border-collapse:separate;
      border-spacing:0;
    }
    .bundle-table th,
    .bundle-table td{
      padding:12px 14px;
      border-bottom:1px solid #edf2f9;
      text-align:left;
      font-size:13px;
      white-space:nowrap;
      vertical-align:middle;
    }
    .bundle-table th{
      position:sticky;
      top:0;
      background:#f5f9ff;
      color:#31507e;
      font-weight:700;
      z-index:1;
    }
    .bundle-table tr:last-child td{
      border-bottom:none;
    }
    .bundle-table .mom{
      font-weight:800;
    }
    .bundle-table .mom.pos{
      color:#ef4444;
    }
    .bundle-table .mom.neg{
      color:#16a34a;
    }
    .bundle-table .mom.flat{
      color:#64748b;
    }
    .bundle-empty{
      padding:18px 12px;
      color:#7b8ba5;
      font-size:13px;
      line-height:1.7;
    }
  </style>
  <script>
    (function(){
      function bundleMomClassFinal(value){
        const n = safeNum(value);
        if(n > 0.005) return "pos";
        if(n < -0.005) return "neg";
        return "flat";
      }
      function ensureBundleLayoutTail(){
        const bundleCard = document.getElementById("bundle-chart")?.closest(".card");
        const bundlePlot = document.getElementById("bundle-chart");
        if(!bundleCard || !bundlePlot) return null;
        bundleCard.classList.add("chart-grid-single-fill","chart-grid-force-wide");
        let leftStack = document.getElementById("bundle-left-stack");
        if(!leftStack){
          leftStack = document.createElement("div");
          leftStack.id = "bundle-left-stack";
          leftStack.className = "bundle-left-stack";
          bundlePlot.parentElement.insertBefore(leftStack, bundlePlot);
          leftStack.appendChild(bundlePlot);
        } else if(bundlePlot.parentElement !== leftStack){
          leftStack.insertBefore(bundlePlot, leftStack.firstChild || null);
        }
        let tableCard = document.getElementById("bundle-table-card");
        if(!tableCard){
          tableCard = document.createElement("div");
          tableCard.id = "bundle-table-card";
          tableCard.className = "bundle-table-card";
          tableCard.innerHTML = '<div class="bundle-table-head"><div class="bundle-table-title">新套组诊断表</div><div class="bundle-table-note">一屏放不下可左右滑动</div></div><div class="bundle-table-wrap" id="bundle-table-wrap"></div>';
          leftStack.appendChild(tableCard);
        }
        return tableCard;
      }
      function renderBundleTableTail(){
        const tableCard = ensureBundleLayoutTail();
        const holder = document.getElementById("bundle-table-wrap");
        if(!tableCard || !holder || typeof buildBundleRows !== "function") return;
        const rows = buildBundleRows();
        if(!rows || !rows.length){
          holder.innerHTML = '<div class="bundle-empty">当前筛选下暂无重点新套组诊断数据。</div>';
          return;
        }
        holder.innerHTML = `<table class="bundle-table"><thead><tr><th>套组</th><th>GMV</th><th>环比</th><th>访客</th><th>环比</th><th>转化率</th><th>环比</th><th>站内花费</th><th>环比</th><th>站外花费</th><th>环比</th><th>站外曝光量</th><th>环比</th></tr></thead><tbody>${rows.map(row=>{
          const innerMom = pctChange(safeNum(row.innerSpend), safeNum(row.priorInnerSpend));
          const outerMom = pctChange(safeNum(row.outerSpend), safeNum(row.priorOuterSpend));
          const impressionMom = pctChange(safeNum(row.impressions), safeNum(row.priorImpressions));
          return `<tr>
            <td>${row.label}</td>
            <td>${fmtMoney(row.gmv)}</td>
            <td><span class="mom ${bundleMomClassFinal(row.gmvMom)}">${fmtDelta(row.gmvMom)}</span></td>
            <td>${fmtNum(row.visitors)}</td>
            <td><span class="mom ${bundleMomClassFinal(row.visitorsMom)}">${fmtDelta(row.visitorsMom)}</span></td>
            <td>${fmtPct(row.cvr)}</td>
            <td><span class="mom ${bundleMomClassFinal(row.cvrMom)}">${fmtDelta(row.cvrMom)}</span></td>
            <td>${fmtMoney(row.innerSpend)}</td>
            <td><span class="mom ${bundleMomClassFinal(innerMom)}">${fmtDelta(innerMom)}</span></td>
            <td>${fmtMoney(row.outerSpend)}</td>
            <td><span class="mom ${bundleMomClassFinal(outerMom)}">${fmtDelta(outerMom)}</span></td>
            <td>${fmtNum(row.impressions)}</td>
            <td><span class="mom ${bundleMomClassFinal(impressionMom)}">${fmtDelta(impressionMom)}</span></td>
          </tr>`;
        }).join("")}</tbody></table>`;
        if(window.Plotly && window.Plotly.Plots && typeof window.Plotly.Plots.resize === "function"){
          try{ window.Plotly.Plots.resize(document.getElementById("bundle-chart")); }catch(err){}
        }
      }
      const prevRenderAllBundleTail = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAllBundleTail === "function") prevRenderAllBundleTail(reset);
        renderBundleTableTail();
      };
      let bundleTailTimer = null;
      window.addEventListener("resize", function(){
        clearTimeout(bundleTailTimer);
        bundleTailTimer = setTimeout(renderBundleTableTail, 120);
      });
      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){ setTimeout(renderBundleTableTail, 0); });
      }else{
        setTimeout(renderBundleTableTail, 0);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      const ROI_KINDS = ["category_roi","product_roi"];
      const ALL_TABLE_KINDS = ["category","product","category_roi","product_roi"];
      const TABLE_NAME_MAP = {
        category: "品类",
        product: "产品",
        category_roi: "品类ROI",
        product_roi: "产品ROI"
      };
      const ROI_COLUMNS = {
        category_roi: [["label","品类"],["gmv","GMV"],["gmvMom","环比"],["fullRoi","全站ROI"],["fullRoiMom","环比"],["adRoi","站外广告ROI"],["adRoiMom","环比"],["seedSpend","种草花费"],["seedSpendMom","环比"],["harvestSpend","收割花费"],["harvestSpendMom","环比"],["diagnosis","诊断"],["todo","下一步"]],
        product_roi: [["label","产品"],["gmv","GMV"],["gmvMom","环比"],["fullRoi","全站ROI"],["fullRoiMom","环比"],["adRoi","站外广告ROI"],["adRoiMom","环比"],["seedSpend","种草花费"],["seedSpendMom","环比"],["harvestSpend","收割花费"],["harvestSpendMom","环比"],["diagnosis","诊断"],["todo","下一步"]]
      };
      const MONEY_KEYS = new Set(["gmv","innerSpend","outerSpend","adGmv","seedSpend","harvestSpend"]);
      const COUNT_KEYS = new Set(["visitors","impressions"]);
      const RATIO_KEYS = new Set(["cvr","gmvMom","visitorsMom","cvrMom","innerSpendMom","outerSpendMom","impressionsMom","fullRoiMom","adRoiMom","seedSpendMom","harvestSpendMom"]);
      const ROI_VALUE_KEYS = new Set(["fullRoi","adRoi"]);
      const MOM_KEYS = new Set(["gmvMom","visitorsMom","cvrMom","innerSpendMom","outerSpendMom","impressionsMom","fullRoiMom","adRoiMom","seedSpendMom","harvestSpendMom"]);
      const BLOCKED_LABELS = new Set(["未归类","鏈綊绫?","闁哄牜浜滅紞濠勭尵?","閺堫亜缍婄猾?","其他","鍏朵粬"]);

      function tableMomClass(value){
        const n = safeNum(value);
        if(n > 0.005) return "mom-pos";
        if(n < -0.005) return "mom-neg";
        return "mom-flat";
      }

      function tableDisplayName(kind){
        return TABLE_NAME_MAP[kind] || "表格";
      }

      function isBlockedLabel(label){
        return BLOCKED_LABELS.has(String(label || "").trim());
      }

      function ensureRoiSections(){
        const specs = [
          {
            afterId: "category-table-controls",
            sectionId: "category-roi-card",
            title: "品类ROI诊断表",
            note: "聚焦全站ROI、站外广告ROI与种草/收割花费结构",
            controlsId: "category_roi-table-controls",
            tableId: "category_roi-table",
          },
          {
            afterId: "product-table-controls",
            sectionId: "product-roi-card",
            title: "产品ROI诊断表",
            note: "判断产品是否还有放量空间，以及更偏种草还是收割",
            controlsId: "product_roi-table-controls",
            tableId: "product_roi-table",
          },
        ];
        specs.forEach((spec) => {
          if(document.getElementById(spec.sectionId)) return;
          const anchorControls = document.getElementById(spec.afterId);
          const anchorSection = anchorControls && anchorControls.closest(".table-card");
          if(!anchorSection || !anchorSection.parentNode) return;
          const section = document.createElement("section");
          section.className = "table-card";
          section.id = spec.sectionId;
          section.style.marginTop = "18px";
          section.innerHTML = `<div class="table-head"><div><div class="section-label">ROI Diagnosis</div><h2>${spec.title}</h2><div class="muted">${spec.note}</div></div></div><div id="${spec.controlsId}"></div><div id="${spec.tableId}"></div>`;
          anchorSection.parentNode.insertBefore(section, anchorSection.nextSibling);
        });
      }

      function ensureRoiTableState(){
        Object.entries(ROI_COLUMNS).forEach(([kind, columns]) => {
          TABLE_COLUMNS[kind] = columns;
          if(!state.tableState[kind]) state.tableState[kind] = initTableState(kind);
        });
      }

      function formatCell(key, value){
        if(MONEY_KEYS.has(key)) return fmtMoney(value);
        if(COUNT_KEYS.has(key)) return fmtNum(value);
        if(ROI_VALUE_KEYS.has(key)) return fmtNum(value, 2);
        if(RATIO_KEYS.has(key)) return fmtPct(value);
        return value ?? "";
      }
      window.formatCell = formatCell;

      function boardBucket(value){
        const text = String(value || "");
        if(text.includes("种草")) return "seed";
        if(text.includes("收割")) return "harvest";
        return "other";
      }

      function buildBoardSpendMap(dim, periodType){
        const rows = aggregateBy(filterRows(DATA.outer_module_daily, periodType), [dim, "board"], ["outer_spend"]);
        const map = new Map();
        rows.forEach((row) => {
          const key = row[dim];
          if(!key || isBlockedLabel(key)) return;
          if(!map.has(key)) map.set(key, { seed: 0, harvest: 0 });
          const bucket = boardBucket(row.board);
          const target = map.get(key);
          if(bucket === "seed") target.seed += safeNum(row.outer_spend);
          if(bucket === "harvest") target.harvest += safeNum(row.outer_spend);
        });
        return map;
      }

      function roiDiagnosisText(row, levelName){
        const direction = row.seedSpend >= row.harvestSpend * 1.2 ? "偏种草" : row.harvestSpend >= row.seedSpend * 1.2 ? "偏收割" : "种收均衡";
        if(row.gmv <= 0 && row.innerSpend + row.outerSpend <= 0){
          return [`${levelName}当前暂无有效投放和成交样本，暂时无法判断放量空间。`,`先确认该${levelName.includes("品类") ? "品类" : "产品"}是否已有稳定投放，再补齐样本后判断种收方向。`];
        }
        if(row.fullRoi >= 4 && row.adRoi >= 1.5){
          return [`${levelName}${direction}，全站ROI ${fmtNum(row.fullRoi,2)}、站外广告ROI ${fmtNum(row.adRoi,2)}，当前还有放量空间。`,`继续保留高ROI素材和人群，若当前偏种草则补强收割链路，若偏收割则稳步加预算吃规模。`];
        }
        if(row.fullRoi >= 4 && row.adRoi < 1){
          return [`${levelName}${direction}，全站ROI 高但站外广告ROI 偏低，更依赖店铺自然承接，放量要克制。`,`先优化站外点击到成交链路，再做小步扩量，避免粗暴加预算稀释效率。`];
        }
        if(row.fullRoi >= 2.5){
          return [`${levelName}${direction}，全站ROI 处于可继续测试区间，但还不算强势放量盘。`,`优先沿高点击率素材和高转化链接补样本，再看是否把预算向更高效方向集中。`];
        }
        if(direction === "偏收割"){
          return [`${levelName}当前偏收割，但全站ROI ${fmtNum(row.fullRoi,2)} 偏低，收割效率不足。`,`收缩低效收割花费，优先更换素材或价格钩子，避免继续硬吃存量人群。`];
        }
        return [`${levelName}当前偏种草，但全站ROI ${fmtNum(row.fullRoi,2)} 偏低，前链路投入重、后链路承接弱。`,`先压缩低效种草预算，把资源回流到高转化商品页和更明确的收割动作上。`];
      }

      function buildRoiDiagnosisRows(kind){
        const dim = kind === "category_roi" ? "category" : "product";
        const currentShop = aggregateBy(filterRows(DATA.shop_daily, "current"), [dim], ["gmv"]);
        const priorShop = aggregateBy(filterRows(DATA.shop_daily, "prior"), [dim], ["gmv"]);
        const currentOuter = aggregateBy(filterRows(DATA.outer_daily, "current"), [dim], ["outer_spend", "ad_gmv"]);
        const priorOuter = aggregateBy(filterRows(DATA.outer_daily, "prior"), [dim], ["outer_spend", "ad_gmv"]);
        const currentInner = aggregateBy(filterRows(DATA.inner_daily, "current"), [dim], ["inner_spend"]);
        const priorInner = aggregateBy(filterRows(DATA.inner_daily, "prior"), [dim], ["inner_spend"]);
        const currentBoard = buildBoardSpendMap(dim, "current");
        const priorBoard = buildBoardSpendMap(dim, "prior");
        const currentShopMap = indexBy(currentShop, dim);
        const priorShopMap = indexBy(priorShop, dim);
        const currentOuterMap = indexBy(currentOuter, dim);
        const priorOuterMap = indexBy(priorOuter, dim);
        const currentInnerMap = indexBy(currentInner, dim);
        const priorInnerMap = indexBy(priorInner, dim);
        const keys = new Set([
          ...currentShopMap.keys(),
          ...priorShopMap.keys(),
          ...currentOuterMap.keys(),
          ...priorOuterMap.keys(),
          ...currentInnerMap.keys(),
          ...priorInnerMap.keys(),
          ...currentBoard.keys(),
          ...priorBoard.keys(),
        ]);
        return Array.from(keys).map((key) => {
          if(!key || isBlockedLabel(key)) return null;
          const currentShopRow = currentShopMap.get(key) || {};
          const priorShopRow = priorShopMap.get(key) || {};
          const currentOuterRow = currentOuterMap.get(key) || {};
          const priorOuterRow = priorOuterMap.get(key) || {};
          const currentInnerRow = currentInnerMap.get(key) || {};
          const priorInnerRow = priorInnerMap.get(key) || {};
          const currentBoardRow = currentBoard.get(key) || { seed: 0, harvest: 0 };
          const priorBoardRow = priorBoard.get(key) || { seed: 0, harvest: 0 };
          const gmv = safeNum(currentShopRow.gmv);
          const priorGmv = safeNum(priorShopRow.gmv);
          const innerSpend = safeNum(currentInnerRow.inner_spend);
          const priorInnerSpend = safeNum(priorInnerRow.inner_spend);
          const outerSpend = safeNum(currentOuterRow.outer_spend);
          const priorOuterSpend = safeNum(priorOuterRow.outer_spend);
          const adGmv = safeNum(currentOuterRow.ad_gmv);
          const priorAdGmv = safeNum(priorOuterRow.ad_gmv);
          const fullRoi = ratio(gmv, innerSpend + outerSpend);
          const priorFullRoi = ratio(priorGmv, priorInnerSpend + priorOuterSpend);
          const adRoi = ratio(adGmv, outerSpend);
          const priorAdRoi = ratio(priorAdGmv, priorOuterSpend);
          const seedSpend = safeNum(currentBoardRow.seed);
          const priorSeedSpend = safeNum(priorBoardRow.seed);
          const harvestSpend = safeNum(currentBoardRow.harvest);
          const priorHarvestSpend = safeNum(priorBoardRow.harvest);
          const [diagnosis, todo] = roiDiagnosisText({
            gmv,
            innerSpend,
            outerSpend,
            fullRoi,
            adRoi,
            seedSpend,
            harvestSpend,
          }, kind === "category_roi" ? `${key}品类` : key);
          return {
            label: key,
            gmv,
            gmvMom: pctChange(gmv, priorGmv),
            fullRoi,
            fullRoiMom: pctChange(fullRoi, priorFullRoi),
            adRoi,
            adRoiMom: pctChange(adRoi, priorAdRoi),
            seedSpend,
            seedSpendMom: pctChange(seedSpend, priorSeedSpend),
            harvestSpend,
            harvestSpendMom: pctChange(harvestSpend, priorHarvestSpend),
            innerSpend,
            outerSpend,
            adGmv,
            diagnosis,
            todo,
          };
        }).filter(Boolean).sort((a,b) => safeNum(b.gmv) - safeNum(a.gmv));
      }

      function renderTable(kind, rows){
        const controls = document.getElementById(`${kind}-table-controls`);
        const holder = document.getElementById(`${kind}-table`);
        if(!controls || !holder || !TABLE_COLUMNS[kind]) return;
        const ts = state.tableState[kind] || (state.tableState[kind] = initTableState(kind));
        const columns = TABLE_COLUMNS[kind];
        const panelOpen = !!ts.panelOpen;
        controls.innerHTML = `<div class="table-toolbar"><button class="btn-secondary" type="button" data-column-toggle="${kind}">列设置</button><button class="btn-secondary" type="button" data-reset-columns="${kind}">恢复所有列</button><button class="btn-secondary" type="button" data-reset="${kind}">恢复所有行</button><button class="btn-secondary" type="button" data-export="${kind}">下载CSV</button><div class="column-panel" id="${kind}-column-panel" ${panelOpen ? "" : "hidden"}><div class="column-grid">${columns.map(([key]) => `<label class="column-item"><input type="checkbox" data-kind="${kind}" data-column="${key}" ${ts.visible[key] ? "checked" : ""}><span>${ts.labels[key]}</span></label>`).join("")}</div></div></div>`;
        const visible = columns.filter(([key]) => ts.visible[key]);
        const list = rows.filter((row) => !ts.hiddenRows.has(row.label)).sort((a,b) => safeNum(b.gmv) - safeNum(a.gmv)).slice(0, state.topN);
        if(!list.length){
          holder.innerHTML = `<div class="empty-state">当前筛选下暂无${tableDisplayName(kind)}诊断表数据。</div>`;
          bindTableActions(kind, rows);
          return;
        }
        holder.innerHTML = `<div class="table-wrap"><table><thead><tr>${visible.map(([key]) => `<th>${ts.labels[key]}</th>`).join("")}<th>操作</th></tr></thead><tbody>${list.map((row) => `<tr>${visible.map(([key]) => {
          const cell = formatCell(key, row[key]);
          return `<td>${MOM_KEYS.has(key) ? `<span class="mom-value ${tableMomClass(row[key])}">${cell}</span>` : cell}</td>`;
        }).join("")}<td><button class="mini-btn" data-hide="${kind}" data-row="${row.label}">隐藏</button></td></tr>`).join("")}</tbody></table></div>`;
        bindTableActions(kind, rows);
      }
      window.renderTable = renderTable;

      function bindTableActions(kind, rows){
        const ts = state.tableState[kind];
        const toggle = document.querySelector(`button[data-column-toggle="${kind}"]`);
        if(toggle) toggle.onclick = (e) => {
          e.stopPropagation();
          ALL_TABLE_KINDS.forEach((name) => {
            if(state.tableState[name]) state.tableState[name].panelOpen = false;
          });
          ts.panelOpen = !ts.panelOpen;
          renderTable(kind, rows);
        };
        const resetColumns = document.querySelector(`button[data-reset-columns="${kind}"]`);
        if(resetColumns) resetColumns.onclick = () => {
          TABLE_COLUMNS[kind].forEach(([key]) => { ts.visible[key] = true; });
          ts.panelOpen = false;
          renderTable(kind, rows);
        };
        document.querySelectorAll(`input[data-kind="${kind}"][data-column]`).forEach((input) => {
          input.onchange = (e) => {
            ts.visible[input.dataset.column] = e.target.checked;
            ts.panelOpen = true;
            renderTable(kind, rows);
          };
        });
        document.querySelectorAll(`button[data-hide="${kind}"]`).forEach((btn) => {
          btn.onclick = () => {
            ts.hiddenRows.add(btn.dataset.row);
            renderTable(kind, rows);
          };
        });
        const resetRows = document.querySelector(`button[data-reset="${kind}"]`);
        if(resetRows) resetRows.onclick = () => {
          ts.hiddenRows.clear();
          renderTable(kind, rows);
        };
        const exportBtn = document.querySelector(`button[data-export="${kind}"]`);
        if(exportBtn) exportBtn.onclick = () => exportCsv(kind, rows);
        const panel = document.getElementById(`${kind}-column-panel`);
        if(panel) panel.onclick = (e) => e.stopPropagation();
      }
      window.bindTableActions = bindTableActions;

      if(!window.__roiTablePanelCloserBound){
        document.addEventListener("click", () => {
          let changed = false;
          ALL_TABLE_KINDS.forEach((kind) => {
            if(state.tableState[kind] && state.tableState[kind].panelOpen){
              state.tableState[kind].panelOpen = false;
              changed = true;
            }
          });
          if(changed){
            renderTable("category", buildDiagnosisRows("category"));
            renderTable("product", buildDiagnosisRows("product"));
            renderTable("category_roi", buildRoiDiagnosisRows("category_roi"));
            renderTable("product_roi", buildRoiDiagnosisRows("product_roi"));
          }
        });
        window.__roiTablePanelCloserBound = true;
      }

      const prevRenderAllRoiTables = window.renderAll;
      window.renderAll = function(reset){
        if(typeof prevRenderAllRoiTables === "function") prevRenderAllRoiTables(reset);
        ensureRoiTableState();
        ensureRoiSections();
        renderTable("category_roi", buildRoiDiagnosisRows("category_roi"));
        renderTable("product_roi", buildRoiDiagnosisRows("product_roi"));
      };

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          ensureRoiTableState();
          ensureRoiSections();
          if(typeof window.renderAll === "function") window.renderAll(true);
        });
      }else{
        ensureRoiTableState();
        ensureRoiSections();
        if(typeof window.renderAll === "function") window.renderAll(true);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .wide-analysis-card{grid-column:1 / -1}
    .analysis-dual-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;margin-top:14px}
    .analysis-stack{display:grid;grid-template-rows:repeat(2,minmax(0,1fr));gap:18px}
    .analysis-block{border:1px solid rgba(46,90,167,.12);border-radius:20px;background:#fff;padding:16px 16px 14px;min-width:0}
    .analysis-block-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:10px}
    .analysis-block-title{font-size:15px;font-weight:800;color:#244883;line-height:1.4}
    .analysis-block-note{margin-top:3px;font-size:12px;line-height:1.5;color:#6b7d98}
    .analysis-download{
      min-height:40px;
      padding:0 16px;
      border-radius:16px;
      border:1px solid rgba(46,90,167,.16);
      background:#ffffff;
      color:var(--blue-deep);
      font-size:12px;
      font-weight:700;
      line-height:1;
      box-shadow:0 6px 18px rgba(46,90,167,.06), inset 0 1px 0 rgba(255,255,255,.9);
      white-space:nowrap;
    }
    .analysis-download:hover{
      background:#f8fbff;
      border-color:rgba(46,90,167,.22);
    }
    .analysis-mini-narrative{margin-top:12px;padding:12px 14px;border-radius:16px;background:#f5f9ff;border:1px solid rgba(46,90,167,.10);font-size:13px;line-height:1.75;color:#334155}
    .analysis-mini-narrative p{margin:0}
    .analysis-mini-narrative p + p{margin-top:6px}
    .analysis-hidden-legacy{display:none !important}
    @media (max-width:1080px){
      .analysis-dual-grid{grid-template-columns:1fr}
    }
  </style>
  <script>
    (function(){
      const prevRenderAllSectionCharts = window.renderAll;
      const HIDDEN_LABELS = new Set(["其他","鍏朵粬","未归类","鏈綊绫?","闁哄牜浜滅紞濠勭尵?","閺堫亜缍婄猾?"]);

      function validLabel(text){
        return !HIDDEN_LABELS.has(String(text || "").trim());
      }

      function bindAnalysisDownloads(scope){
        if(!scope) return;
        scope.querySelectorAll("[data-plot-download]").forEach((btn)=>{
          btn.onclick = ()=>downloadCanvas(btn.dataset.plotDownload);
        });
      }

      function mountSectionCards(){
        const moduleCard = document.getElementById("module-chart")?.closest(".card");
        if(moduleCard && !moduleCard.dataset.analysisMounted){
          moduleCard.dataset.analysisMounted = "1";
          moduleCard.classList.add("wide-analysis-card");
          moduleCard.innerHTML = `
            <div class="analysis-hidden-legacy"><div id="module-chart"></div><div id="module-narrative"></div></div>
            <div class="chart-head"><div><div class="section-label">Module</div><h2>模块分析</h2><div class="muted">全部使用站外广告口径，聚焦广告GMV、CTR、CVR、CPM与曝光量</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-stack">
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">广告GMV柱状 + 转化率曲线</div><div class="analysis-block-note">模块维度看当前成交规模与点击后转化效率</div></div><button class="analysis-download" data-plot-download="module-gmv-cvr-chart">下载图表</button></div>
                  <div id="module-gmv-cvr-chart" class="plot"></div>
                  <div id="module-gmv-cvr-narrative" class="analysis-mini-narrative"></div>
                </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">广告GMV柱状 + CTR曲线</div><div class="analysis-block-note">模块维度仅看站外广告口径下的规模与点击吸引力</div></div><button class="analysis-download" data-plot-download="module-gmv-ctr-chart">下载图表</button></div>
                  <div id="module-gmv-ctr-chart" class="plot"></div>
                  <div id="module-gmv-ctr-narrative" class="analysis-mini-narrative"></div>
                </div>
              </div>
              <div class="analysis-stack">
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">CPM曲线</div><div class="analysis-block-note">看各模块拿量成本高低，识别高成本拖累点</div></div><button class="analysis-download" data-plot-download="module-cpm-chart">下载图表</button></div>
                  <div id="module-cpm-chart" class="plot"></div>
                  <div id="module-cpm-narrative" class="analysis-mini-narrative"></div>
                </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">曝光量曲线</div><div class="analysis-block-note">看模块拿量规模，判断预算是否过度集中</div></div><button class="analysis-download" data-plot-download="module-impressions-chart">下载图表</button></div>
                  <div id="module-impressions-chart" class="plot"></div>
                  <div id="module-impressions-narrative" class="analysis-mini-narrative"></div>
                </div>
              </div>
            </div>`;
          bindAnalysisDownloads(moduleCard);
        }

        const categoryCard = document.getElementById("category-chart")?.closest(".card");
        if(categoryCard && !categoryCard.dataset.analysisMounted){
          categoryCard.dataset.analysisMounted = "1";
          categoryCard.classList.add("wide-analysis-card");
          categoryCard.innerHTML = `
            <div class="analysis-hidden-legacy"><div id="category-chart"></div><div id="category-narrative"></div></div>
            <div class="chart-head"><div><div class="section-label">Category</div><h2>品类分析</h2><div class="muted">从GMV、转化率与访客结构看品类增长弹性</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">GMV柱状 + 转化率曲线</div><div class="analysis-block-note">看高体量品类里，哪些已经兼顾成交规模和转化效率</div></div><button class="analysis-download" data-plot-download="category-gmv-cvr-chart">下载图表</button></div>
                <div id="category-gmv-cvr-chart" class="plot"></div>
                <div id="category-gmv-cvr-narrative" class="analysis-mini-narrative"></div>
              </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">站内外花费柱状 + 访客曲线</div><div class="analysis-block-note">看品类拿量投入与访客承接是否匹配，并对比上月同期</div></div><button class="analysis-download" data-plot-download="category-gmv-visitors-chart">下载图表</button></div>
                <div id="category-gmv-visitors-chart" class="plot"></div>
                <div id="category-gmv-visitors-narrative" class="analysis-mini-narrative"></div>
              </div>
            </div>`;
          bindAnalysisDownloads(categoryCard);
        }

        const productCard = document.getElementById("product-chart")?.closest(".card");
        if(productCard && !productCard.dataset.analysisMounted){
          productCard.dataset.analysisMounted = "1";
          productCard.classList.add("wide-analysis-card");
          productCard.innerHTML = `
            <div class="analysis-hidden-legacy"><div id="product-chart"></div><div id="product-narrative"></div></div>
            <div class="chart-head"><div><div class="section-label">Product</div><h2>产品分析</h2><div class="muted">看主力产品是靠高转化放大，还是靠大流量硬推规模</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">GMV柱状 + 转化率曲线</div><div class="analysis-block-note">识别当前既能跑规模、又能跑效率的主力产品</div></div><button class="analysis-download" data-plot-download="product-gmv-cvr-chart">下载图表</button></div>
                <div id="product-gmv-cvr-chart" class="plot"></div>
                <div id="product-gmv-cvr-narrative" class="analysis-mini-narrative"></div>
              </div>
                <div class="analysis-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">站内外花费柱状 + 访客曲线</div><div class="analysis-block-note">看产品层的投放投入与访客承接是否匹配，并对比上月同期</div></div><button class="analysis-download" data-plot-download="product-gmv-visitors-chart">下载图表</button></div>
                <div id="product-gmv-visitors-chart" class="plot"></div>
                <div id="product-gmv-visitors-narrative" class="analysis-mini-narrative"></div>
              </div>
            </div>`;
          bindAnalysisDownloads(productCard);
        }

        const ttCard = document.getElementById("tt-gap-chart")?.closest(".card");
        if(ttCard && !ttCard.dataset.analysisMounted){
          ttCard.dataset.analysisMounted = "1";
          ttCard.classList.add("wide-analysis-card");
          ttCard.innerHTML = `
            <div class="analysis-hidden-legacy"><div id="tt-gap-chart"></div><div id="tt-gap-narrative"></div></div>
            <div class="chart-head"><div><div class="section-label">TT Gap</div><h2>TT销量GAP</h2><div class="muted">同时看Shopee与TT销量规模，以及产品维度的销量缺口压力</div></div></div>
            <div class="analysis-dual-grid">
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">Shopee销量 vs TT销量</div><div class="analysis-block-note">看各主力产品在Shopee与TT的销量体量差异</div></div><button class="analysis-download" data-plot-download="tt-gap-compare-chart">下载图表</button></div>
                <div id="tt-gap-compare-chart" class="plot"></div>
                <div id="tt-gap-compare-narrative" class="analysis-mini-narrative"></div>
              </div>
              <div class="analysis-block">
                <div class="analysis-block-head"><div><div class="analysis-block-title">销量GAP柱状 + 达成率曲线</div><div class="analysis-block-note">柱看Shopee相对TT的绝对差值，线看Shopee/TT达成率</div></div><button class="analysis-download" data-plot-download="tt-gap-delta-chart">下载图表</button></div>
                <div id="tt-gap-delta-chart" class="plot"></div>
                <div id="tt-gap-delta-narrative" class="analysis-mini-narrative"></div>
              </div>
            </div>`;
          bindAnalysisDownloads(ttCard);
        }
      }

      function axisCategoryLabels(rows){
        return rows.map((row)=>shortLabel(row.label, 12));
      }

      function sortByGmv(rows){
        return [...rows].filter((row)=>row && validLabel(row.label)).sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      }

      function drawBarLineCombo(id, rows, options){
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          const el = document.getElementById(id);
          if(el) el.innerHTML = '<div class="empty-state">当前筛选下暂无图表数据。</div>';
          return;
        }
        const labels = axisCategoryLabels(list);
        const barTickFormat = options.barIsCount ? ",.0f" : ",.0f";
        const lineTickFormat = options.lineIsPercent ? ".0%" : options.lineIsMoney ? ",.0f" : ",.0f";
        renderPlot(id,[
          {
            type:"bar",
            name:options.barName,
            x:labels,
            y:list.map((row)=>safeNum(row[options.barField])),
            customdata:list.map((row)=>row.label),
            marker:{color:options.barColor || "#4f8df7",line:{width:0}},
            hovertemplate:`<b>%{customdata}</b><br>${options.barName}: %{y:,.0f}<extra></extra>`
          },
          {
            type:"bar",
            name:options.priorBarName || `上月同期${options.barName}`,
            x:labels,
            y:list.map((row)=>safeNum(row[options.priorBarField || options.barField])),
            customdata:list.map((row)=>row.label),
            marker:{color:options.priorBarColor || "#b7d5ff",line:{width:0}},
            hovertemplate:`<b>%{customdata}</b><br>${options.priorBarName || `上月同期${options.barName}`}: %{y:,.0f}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:options.lineName,
            x:labels,
            y:list.map((row)=>safeNum(row[options.lineField])),
            customdata:list.map((row)=>row.label),
            yaxis:"y2",
            line:{color:options.lineColor || "#f59e0b",width:3},
            marker:{size:7,color:options.lineColor || "#f59e0b"},
            hovertemplate:options.lineHover || `<b>%{customdata}</b><br>${options.lineName}: ${options.lineIsPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:options.priorLineName || `上月同期${options.lineName}`,
            x:labels,
            y:list.map((row)=>safeNum(row[options.priorLineField || options.lineField])),
            customdata:list.map((row)=>row.label),
            yaxis:"y2",
            line:{color:options.priorLineColor || "#fcd34d",width:3},
            marker:{size:7,color:options.priorLineColor || "#fcd34d"},
            hovertemplate:options.priorLineHover || `<b>%{customdata}</b><br>${options.priorLineName || `上月同期${options.lineName}`}: ${options.lineIsPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          }
        ],{
          height:360,
          barmode:"group",
          bargap:.22,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:66,r:70,t:28,b:116},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-26,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:options.barName,gridcolor:"#eef3fb",zeroline:false,tickformat:barTickFormat,tickfont:{color:"#64748b"}},
          yaxis2:{title:options.lineName,overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:lineTickFormat,tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function drawSpendVisitorsCombo(id, rows){
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          const el = document.getElementById(id);
          if(el) el.innerHTML = '<div class="empty-state">褰撳墠绛涢€変笅鏆傛棤鍥捐〃鏁版嵁銆?/div>';
          return;
        }
        const labels = axisCategoryLabels(list);
        renderPlot(id,[
          {type:"bar",name:"站内广告费",x:labels,y:list.map((row)=>safeNum(row.innerSpend)),customdata:list.map((row)=>row.label),marker:{color:"#2E5AA7",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>站内广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"上月同期站内广告费",x:labels,y:list.map((row)=>safeNum(row.priorInnerSpend)),customdata:list.map((row)=>row.label),marker:{color:"#B8DCFF",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>上月同期站内广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"站外广告费",x:labels,y:list.map((row)=>safeNum(row.spend)),customdata:list.map((row)=>row.label),marker:{color:"#FFA62B",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>站外广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"上月同期站外广告费",x:labels,y:list.map((row)=>safeNum(row.priorSpend)),customdata:list.map((row)=>row.label),marker:{color:"#F8E6A0",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>上月同期站外广告费: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"CTR",x:labels,y:list.map((row)=>safeNum(row.ctr)),customdata:list.map((row)=>row.label),yaxis:"y2",line:{color:"#22c55e",width:3},marker:{size:7,color:"#22c55e"},hovertemplate:"<b>%{customdata}</b><br>CTR: %{y:.1%}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"上月同期CTR",x:labels,y:list.map((row)=>safeNum(row.priorCtr)),customdata:list.map((row)=>row.label),yaxis:"y2",line:{color:"#86efac",width:3},marker:{size:7,color:"#86efac"},hovertemplate:"<b>%{customdata}</b><br>上月同期CTR: %{y:.1%}<extra></extra>"}
        ],{
          height:376,
          barmode:"group",
          bargap:.18,
          bargroupgap:.08,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:66,r:74,t:28,b:138},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:-0.28,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-26,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"广告费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:"CTR",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function drawDualLine(id, rows, options){
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          const el = document.getElementById(id);
          if(el) el.innerHTML = '<div class="empty-state">当前筛选下暂无图表数据。</div>';
          return;
        }
        const labels = axisCategoryLabels(list);
        renderPlot(id,[
          {
            type:"scatter",
            mode:"lines+markers",
            name:options.name,
            x:labels,
            y:list.map((row)=>safeNum(row[options.field])),
            customdata:list.map((row)=>row.label),
            line:{color:options.color || "#2f6bff",width:3},
            marker:{size:7,color:options.color || "#2f6bff"},
            hovertemplate:options.hover || `<b>%{customdata}</b><br>${options.name}: ${options.isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:options.priorName || `上月同期${options.name}`,
            x:labels,
            y:list.map((row)=>safeNum(row[options.priorField || options.field])),
            customdata:list.map((row)=>row.label),
            line:{color:options.priorColor || "#9ec6ff",width:3},
            marker:{size:7,color:options.priorColor || "#9ec6ff"},
            hovertemplate:options.priorHover || `<b>%{customdata}</b><br>${options.priorName || `上月同期${options.name}`}: ${options.isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          }
        ],{
          height:300,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:62,r:20,t:28,b:110},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.16,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-26,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:options.name,gridcolor:"#eef3fb",zeroline:false,tickformat:options.isPercent ? ".0%" : ",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function buildModuleInnerSpendMap(periodType){
        const outerRows = filterRows(DATA.outer_module_daily, periodType).filter((row)=>row && validLabel(row.module));
        const grouped = new Map();
        outerRows.forEach((row)=>{
          const key = [row.day,row.shop,row.group,row.category,row.product].join("||");
          if(!grouped.has(key)) grouped.set(key, []);
          grouped.get(key).push(row);
        });
        const moduleSpendMap = new Map();
        filterRows(DATA.inner_daily, periodType).forEach((row)=>{
          const key = [row.day,row.shop,row.group,row.category,row.product].join("||");
          const matched = grouped.get(key);
          if(!matched || !matched.length) return;
          const innerSpend = safeNum(row.inner_spend);
          const totalOuterSpend = matched.reduce((sum, item)=>sum + safeNum(item.outer_spend), 0);
          matched.forEach((item)=>{
            const share = totalOuterSpend > 0 ? safeNum(item.outer_spend) / totalOuterSpend : 1 / matched.length;
            const currentValue = safeNum(moduleSpendMap.get(item.module));
            moduleSpendMap.set(item.module, currentValue + innerSpend * share);
          });
        });
        return moduleSpendMap;
      }

      function buildModuleOuterRows(){
        const current = aggregateBy(filterRows(DATA.outer_module_daily,"current"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const prior = aggregateBy(filterRows(DATA.outer_module_daily,"prior"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const currentMap = indexBy(current, "module");
        const priorMap = indexBy(prior, "module");
        const currentInnerMap = buildModuleInnerSpendMap("current");
        const priorInnerMap = buildModuleInnerSpendMap("prior");
        const keys = new Set([
          ...currentMap.keys(),
          ...priorMap.keys(),
          ...currentInnerMap.keys(),
          ...priorInnerMap.keys()
        ]);
        return Array.from(keys)
          .map((key)=>{
            const row = currentMap.get(key) || {};
            const priorRow = priorMap.get(key) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            const clicks = safeNum(row.clicks);
            const orders = safeNum(row.ad_orders);
            const priorSpend = safeNum(priorRow.outer_spend);
            const priorImpressions = safeNum(priorRow.impressions);
            const priorClicks = safeNum(priorRow.clicks);
            const priorOrders = safeNum(priorRow.ad_orders);
            return {
              label: key,
              gmv: safeNum(row.ad_gmv),
              priorGmv: safeNum(priorRow.ad_gmv),
              cvr: ratio(orders, clicks),
              priorCvr: ratio(priorOrders, priorClicks),
              ctr: ratio(clicks, impressions),
              priorCtr: ratio(priorClicks, priorImpressions),
              cpm: impressions ? spend / impressions * 1000 : 0,
              priorCpm: priorImpressions ? priorSpend / priorImpressions * 1000 : 0,
              impressions,
              priorImpressions,
              spend,
              priorSpend,
              innerSpend: safeNum(currentInnerMap.get(key)),
              priorInnerSpend: safeNum(priorInnerMap.get(key)),
              clicks,
              priorClicks,
              orders,
              priorOrders
            };
          })
          .filter((row)=>row.label && validLabel(row.label))
          .sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      }

      function drawSpendVisitorsComboV2(id, rows){
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          const el = document.getElementById(id);
          if(el) el.innerHTML = '<div class="empty-state">褰撳墠绛涢€変笅鏆傛棤鍥捐〃鏁版嵁銆?/div>';
          return;
        }
        const labels = axisCategoryLabels(list);
        renderPlot(id,[
          {type:"bar",name:"站内广告费",x:labels,y:list.map((row)=>safeNum(row.innerSpend)),customdata:list.map((row)=>row.label),marker:{color:"#2E5AA7",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>站内广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"上月同期站内广告费",x:labels,y:list.map((row)=>safeNum(row.priorInnerSpend)),customdata:list.map((row)=>row.label),marker:{color:"#B8DCFF",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>上月同期站内广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"站外广告费",x:labels,y:list.map((row)=>safeNum(row.outerSpend)),customdata:list.map((row)=>row.label),marker:{color:"#FFA62B",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>站外广告费: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"上月同期站外广告费",x:labels,y:list.map((row)=>safeNum(row.priorOuterSpend)),customdata:list.map((row)=>row.label),marker:{color:"#F8E6A0",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>上月同期站外广告费: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"访客",x:labels,y:list.map((row)=>safeNum(row.visitors)),customdata:list.map((row)=>row.label),yaxis:"y2",line:{color:"#22c55e",width:3},marker:{size:7,color:"#22c55e"},hovertemplate:"<b>%{customdata}</b><br>访客: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"上月同期访客",x:labels,y:list.map((row)=>safeNum(row.priorVisitors)),customdata:list.map((row)=>row.label),yaxis:"y2",line:{color:"#86efac",width:3},marker:{size:7,color:"#86efac"},hovertemplate:"<b>%{customdata}</b><br>上月同期访客: %{y:,.0f}<extra></extra>"}
        ],{
          height:376,
          barmode:"group",
          bargap:.18,
          bargroupgap:.08,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:66,r:74,t:28,b:138},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:-0.28,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-26,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"广告费",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:"访客",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function buildModuleOuterRowsPure(){
        const current = aggregateBy(filterRows(DATA.outer_module_daily,"current"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const prior = aggregateBy(filterRows(DATA.outer_module_daily,"prior"),["module"],["outer_spend","impressions","clicks","ad_gmv","ad_orders"]);
        const currentMap = indexBy(current, "module");
        const priorMap = indexBy(prior, "module");
        const keys = new Set([...currentMap.keys(), ...priorMap.keys()]);
        return Array.from(keys)
          .map((key)=>{
            const row = currentMap.get(key) || {};
            const priorRow = priorMap.get(key) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            const clicks = safeNum(row.clicks);
            const orders = safeNum(row.ad_orders);
            const priorSpend = safeNum(priorRow.outer_spend);
            const priorImpressions = safeNum(priorRow.impressions);
            const priorClicks = safeNum(priorRow.clicks);
            const priorOrders = safeNum(priorRow.ad_orders);
            return {
              label: key,
              gmv: safeNum(row.ad_gmv),
              priorGmv: safeNum(priorRow.ad_gmv),
              cvr: ratio(orders, clicks),
              priorCvr: ratio(priorOrders, priorClicks),
              ctr: ratio(clicks, impressions),
              priorCtr: ratio(priorClicks, priorImpressions),
              cpm: impressions ? spend / impressions * 1000 : 0,
              priorCpm: priorImpressions ? priorSpend / priorImpressions * 1000 : 0,
              impressions,
              priorImpressions,
              spend,
              priorSpend,
              clicks,
              priorClicks,
              orders,
              priorOrders
            };
          })
          .filter((row)=>row.label && validLabel(row.label))
          .sort((a,b)=>safeNum(b.gmv)-safeNum(a.gmv));
      }

      function renderCategorySection(){
        const rows = sortByGmv(buildDimensionRows("category"));
        drawBarLineCombo("category-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"GMV",
          priorBarName:"上月同期GMV",
          lineName:"转化率",
          priorLineName:"上月同期转化率",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#f59e0b",
          priorLineColor:"#fcd34d",
          lineIsPercent:true,
        });
        drawBarLineCombo("category-gmv-visitors-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"visitors",
          priorLineField:"priorVisitors",
          barName:"GMV",
          priorBarName:"上月同期GMV",
          lineName:"访客",
          priorLineName:"上月同期访客",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#22c55e",
          priorLineColor:"#86efac",
        });
        drawSpendVisitorsComboV2("category-gmv-visitors-chart", rows);
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.visitors)>0).sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const topVisitor = [...rows].sort((a,b)=>safeNum(b.visitors)-safeNum(a.visitors))[0];
        const categoryCvrNarrative = document.getElementById("category-gmv-cvr-narrative");
        if(categoryCvrNarrative){
          categoryCvrNarrative.innerHTML = topGmv ? `<p>${topGmv.label} 当前GMV最高，为 ${fmtMoney(topGmv.gmv)}。</p><p>${bestCvr ? `${bestCvr.label} 转化率最高，为 ${fmtPct(bestCvr.cvr)}，更适合继续承接放量。` : "当前筛选下暂无有效转化率样本。"} 下一步：优先把预算和资源向高GMV且转化率不弱的品类集中。</p>` : '<p>当前筛选下暂无品类数据。</p>';
        }
        const categoryVisitorNarrative = document.getElementById("category-gmv-visitors-narrative");
        if(categoryVisitorNarrative){
          categoryVisitorNarrative.innerHTML = topVisitor ? `<p>${topVisitor.label} 当前访客最高，为 ${fmtNum(topVisitor.visitors)}，说明它是主要流量承接入口。</p><p>${topGmv && topVisitor && topGmv.label !== topVisitor.label ? `${topGmv.label} 虽然GMV领先，但流量并非最高，说明更多靠转化效率在带规模。` : "流量与成交规模集中在同一主力品类，后续要防止预算进一步扎堆。"} 下一步：对高访客低GMV品类，优先回查价格力和详情页承接。</p>` : '<p>当前筛选下暂无品类流量数据。</p>';
        }
      }

      function renderProductSection(){
        const rows = sortByGmv(buildDimensionRows("product"));
        drawBarLineCombo("product-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"GMV",
          priorBarName:"上月同期GMV",
          lineName:"转化率",
          priorLineName:"上月同期转化率",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#f59e0b",
          priorLineColor:"#fcd34d",
          lineIsPercent:true,
        });
        drawBarLineCombo("product-gmv-visitors-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"visitors",
          priorLineField:"priorVisitors",
          barName:"GMV",
          priorBarName:"上月同期GMV",
          lineName:"访客",
          priorLineName:"上月同期访客",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#22c55e",
          priorLineColor:"#86efac",
        });
        drawSpendVisitorsComboV2("product-gmv-visitors-chart", rows);
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.visitors)>0).sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const topVisitor = [...rows].sort((a,b)=>safeNum(b.visitors)-safeNum(a.visitors))[0];
        const productCvrNarrative = document.getElementById("product-gmv-cvr-narrative");
        if(productCvrNarrative){
          productCvrNarrative.innerHTML = topGmv ? `<p>${topGmv.label} 当前GMV最高，为 ${fmtMoney(topGmv.gmv)}。</p><p>${bestCvr ? `${bestCvr.label} 转化率最高，为 ${fmtPct(bestCvr.cvr)}，说明成交效率更强。` : "当前筛选下暂无有效转化率样本。"} 下一步：优先加码高GMV且CVR稳定的产品，避免继续给低转化SKU硬推流量。</p>` : '<p>当前筛选下暂无产品数据。</p>';
        }
        const productVisitorNarrative = document.getElementById("product-gmv-visitors-narrative");
        if(productVisitorNarrative){
          productVisitorNarrative.innerHTML = topVisitor ? `<p>${topVisitor.label} 当前访客最高，为 ${fmtNum(topVisitor.visitors)}。</p><p>${topGmv && topVisitor && topGmv.label !== topVisitor.label ? `${topGmv.label} 的GMV高于其访客排名，说明它更多依赖转化承接而不是纯流量堆量。` : "高流量产品同时也是高GMV产品，后续更要监控素材疲劳和转化回落。"} 下一步：对高访客低GMV产品，优先优化素材承诺和详情页转化链路。</p>` : '<p>当前筛选下暂无产品流量数据。</p>';
        }
      }

      function renderModuleSection(){
        const rows = buildModuleOuterRows();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"广告GMV",
          priorBarName:"上月同期广告GMV",
          lineName:"转化率",
          priorLineName:"上月同期转化率",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#f59e0b",
          priorLineColor:"#fcd34d",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"广告GMV",
          priorBarName:"上月同期广告GMV",
          lineName:"CTR",
          priorLineName:"上月同期CTR",
          barColor:"#4f8df7",
          priorBarColor:"#b7d5ff",
          lineColor:"#22c55e",
          priorLineColor:"#86efac",
          lineIsPercent:true,
        });
        drawDualLine("module-cpm-chart", rows, {
          field:"cpm",
          priorField:"priorCpm",
          name:"CPM",
          color:"#2f6bff",
          priorName:"上月同期CPM",
          priorColor:"#9ec6ff",
          hover:"<b>%{customdata}</b><br>CPM: %{y:,.0f}<extra></extra>",
          priorHover:"<b>%{customdata}</b><br>上月同期CPM: %{y:,.0f}<extra></extra>"
        });
        drawDualLine("module-impressions-chart", rows, {
          field:"impressions",
          priorField:"priorImpressions",
          name:"曝光量",
          color:"#38bdf8",
          priorName:"上月同期曝光量",
          priorColor:"#bae6fd",
          hover:"<b>%{customdata}</b><br>曝光量: %{y:,.0f}<extra></extra>",
          priorHover:"<b>%{customdata}</b><br>上月同期曝光量: %{y:,.0f}<extra></extra>"
        });
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks)>0).sort((a,b)=>safeNum(b.cvr)-safeNum(a.cvr))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions)>0).sort((a,b)=>safeNum(b.ctr)-safeNum(a.ctr))[0];
        const worstCpm = [...rows].filter((row)=>safeNum(row.impressions)>0).sort((a,b)=>safeNum(b.cpm)-safeNum(a.cpm))[0];
        const topImpressions = [...rows].sort((a,b)=>safeNum(b.impressions)-safeNum(a.impressions))[0];
        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          cvrNarrative.innerHTML = topGmv ? `<p>${topGmv.label} 当前广告GMV最高，为 ${fmtMoney(topGmv.gmv)}。</p><p>${bestCvr ? `${bestCvr.label} 转化率最高，为 ${fmtPct(bestCvr.cvr)}，说明点击后成交效率更强。` : "当前筛选下暂无有效模块转化样本。"} 下一步：优先保留高GMV且CVR不弱的模块预算。</p>` : '<p>当前筛选下暂无模块数据。</p>';
        }
        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          ctrNarrative.innerHTML = bestCtr ? `<p>${bestCtr.label} CTR最高，为 ${fmtPct(bestCtr.ctr)}，素材点击吸引力更强。</p><p>${topGmv && bestCtr && topGmv.label !== bestCtr.label ? `${topGmv.label} 虽然广告GMV领先，但点击吸引力不一定最强，后续要继续优化素材首屏。` : "当前主力模块同时兼顾规模和点击率，可继续做稳态放量。"} 下一步：高GMV但CTR偏弱的模块优先换素材和前3秒表达。</p>` : '<p>当前筛选下暂无模块点击率样本。</p>';
        }
        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          cpmNarrative.innerHTML = worstCpm ? `<p>${worstCpm.label} 当前CPM最高，为 ${fmtNum(worstCpm.cpm,2)}，拿量成本压力更大。</p><p>若该模块同时没有跑出相应广告GMV，就会成为明显的投放拖累。 下一步：先回查高CPM模块的人群过窄、素材疲劳或版位问题。</p>` : '<p>当前筛选下暂无模块CPM样本。</p>';
        }
        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          impressionsNarrative.innerHTML = topImpressions ? `<p>${topImpressions.label} 当前曝光量最高，为 ${fmtNum(topImpressions.impressions)}。</p><p>${topGmv && topImpressions && topGmv.label !== topImpressions.label ? `${topImpressions.label} 曝光最多但未必产出最高，说明规模不等于成交。` : "曝光与广告GMV集中在同一模块，后续要关注预算过度集中带来的边际递减。"} 下一步：优先核查高曝光低广告GMV模块的点击和转化衔接。</p>` : '<p>当前筛选下暂无模块曝光样本。</p>';
        }
      }

      function renderTTSection(){
        const rows = buildTTGapRows().filter((row)=>validLabel(row.label)).slice(0, Math.max(1, state.topN));
        if(!rows.length){
          ["tt-gap-compare-chart","tt-gap-delta-chart"].forEach((id)=>{
            const el = document.getElementById(id);
            if(el) el.innerHTML = '<div class="empty-state">当前筛选下暂无TT销量数据。</div>';
          });
          return;
        }
        const labels = rows.map((row)=>shortLabel(row.label, 10));
        renderPlot("tt-gap-compare-chart",[
          {type:"bar",name:"Shopee销量",x:labels,y:rows.map((row)=>safeNum(row.shopee)),customdata:rows.map((row)=>row.label),marker:{color:"#4f8df7",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>Shopee销量: %{y:,.0f}<extra></extra>"},
          {type:"bar",name:"TT销量",x:labels,y:rows.map((row)=>safeNum(row.tt)),customdata:rows.map((row)=>row.label),marker:{color:"#b7d5ff",line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>TT销量: %{y:,.0f}<extra></extra>"}
        ],{
          height:360,
          barmode:"group",
          bargap:.22,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:64,r:24,t:28,b:116},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-24,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"销量",gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
        renderPlot("tt-gap-delta-chart",[
          {type:"bar",name:"销量GAP",x:labels,y:rows.map((row)=>safeNum(row.gap)),customdata:rows.map((row)=>row.label),marker:{color:rows.map((row)=>safeNum(row.gap)>=0?"#4f8df7":"#f97316"),line:{width:0}},hovertemplate:"<b>%{customdata}</b><br>销量GAP: %{y:,.0f}<extra></extra>"},
          {type:"scatter",mode:"lines+markers",name:"Shopee/TT达成率",x:labels,y:rows.map((row)=>safeNum(row.tt)>0?safeNum(row.shopee)/safeNum(row.tt):0),customdata:rows.map((row)=>row.label),yaxis:"y2",line:{color:"#22c55e",width:3},marker:{size:7,color:"#22c55e"},hovertemplate:"<b>%{customdata}</b><br>Shopee/TT达成率: %{y:.1%}<extra></extra>"}
        ],{
          height:360,
          barmode:"group",
          bargap:.22,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:64,r:70,t:28,b:116},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.18,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:-24,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{title:"销量GAP",gridcolor:"#eef3fb",zeroline:true,zerolinecolor:"#d6e2f5",tickformat:",.0f",tickfont:{color:"#64748b"}},
          yaxis2:{title:"Shopee/TT达成率",overlaying:"y",side:"right",showgrid:false,zeroline:false,tickformat:".0%",tickfont:{color:"#64748b"}}
        },basePlotConfig());
        const gapLeader = [...rows].sort((a,b)=>safeNum(b.gap)-safeNum(a.gap))[0];
        const gapLagger = [...rows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];
        const bestRate = [...rows].filter((row)=>safeNum(row.tt)>0).sort((a,b)=>(safeNum(b.shopee)/Math.max(safeNum(b.tt),1))-(safeNum(a.shopee)/Math.max(safeNum(a.tt),1)))[0];
        const compareNarrative = document.getElementById("tt-gap-compare-narrative");
        if(compareNarrative){
          compareNarrative.innerHTML = gapLeader ? `<p>${gapLeader.label} 当前Shopee相对TT领先最多，GAP为 ${fmtNum(gapLeader.gap)}。</p><p>${gapLagger ? `${gapLagger.label} 是当前主要补差压力点，Shopee较TT落后 ${fmtNum(Math.abs(gapLagger.gap))}。` : "当前筛选下暂无明显落后产品。"} 下一步：优先针对落后产品补素材、补站内承接和价格钩子。</p>` : '<p>当前筛选下暂无TT GAP样本。</p>';
        }
        const deltaNarrative = document.getElementById("tt-gap-delta-narrative");
        if(deltaNarrative){
          deltaNarrative.innerHTML = bestRate ? `<p>${bestRate.label} 的Shopee/TT达成率最高，为 ${fmtPct(safeNum(bestRate.tt)>0?safeNum(bestRate.shopee)/safeNum(bestRate.tt):0)}。</p><p>${gapLagger ? `${gapLagger.label} 当前GAP最深，说明它更需要做专项补差。` : "当前筛选下各主力产品与TT差距相对可控。"} 下一步：把预算和资源优先向GAP为负且达成率偏低的产品倾斜。</p>` : '<p>当前筛选下暂无TT达成率样本。</p>';
        }
      }

      window.renderAll = function(reset){
        if(typeof prevRenderAllSectionCharts === "function") prevRenderAllSectionCharts(reset);
        mountSectionCards();
        renderCategorySection();
        renderProductSection();
        renderModuleSection();
        renderTTSection();
      };

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          mountSectionCards();
          if(typeof window.renderAll === "function") window.renderAll(true);
        });
      }else{
        mountSectionCards();
        if(typeof window.renderAll === "function") window.renderAll(true);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .tt-gap-wide-card{grid-column:1 / -1}
    .tt-gap-layout{display:grid;grid-template-columns:minmax(0,1.42fr) minmax(0,.98fr);gap:18px;margin-top:14px;align-items:start;width:100%}
    .tt-gap-table-panel,.tt-gap-chart-panel{min-width:0;border:1px solid rgba(46,90,167,.12);border-radius:22px;background:#fff;padding:16px}
    .tt-gap-table-head,.tt-gap-chart-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px}
    .tt-gap-table-title,.tt-gap-chart-title{font-size:15px;font-weight:800;color:#244883;line-height:1.4}
    .tt-gap-table-wrap{overflow:auto;border-radius:16px;border:1px solid rgba(148,163,184,.14);background:#fff}
    .tt-gap-table{width:100%;min-width:860px;border-collapse:separate;border-spacing:0}
    .tt-gap-table th,.tt-gap-table td{padding:10px 12px;border-bottom:1px solid #edf2f9;text-align:left;font-size:12px;white-space:nowrap;vertical-align:middle}
    .tt-gap-table th{position:sticky;top:0;background:#f5f9ff;color:#31507e;font-weight:700;z-index:1}
    .tt-gap-table tr:last-child td{border-bottom:none}
    .tt-rank{width:38px;color:#6b7d98}
    .tt-gap-value.pos{color:#2f6bff;font-weight:800}
    .tt-gap-value.neg{color:#f97316;font-weight:800}
    .tt-gap-chart-stack{display:grid;grid-template-rows:repeat(2,minmax(0,1fr));gap:18px}
    .tt-gap-chart-block{border:1px solid rgba(46,90,167,.1);border-radius:18px;background:#f8fbff;padding:14px}
    .tt-gap-chart-note{margin-top:10px;padding:12px 14px;border-radius:16px;background:#f5f9ff;border:1px solid rgba(46,90,167,.10);font-size:13px;line-height:1.75;color:#334155}
    @media (max-width:1180px){
      .tt-gap-layout{grid-template-columns:1fr}
      .tt-gap-table{min-width:720px}
    }
  </style>
  <script>
    (function(){
      const prevRenderAllTTGapDetail = window.renderAll;
      const TT_HIDDEN_LABELS = new Set(["其他","鍏朵粬","未归类","鏈綊绫?","闁哄牜浜滅紞濠勭尵?","閺堫亜缍婄猾?"]);

      function ttValidLabel(text){
        return !TT_HIDDEN_LABELS.has(String(text || "").trim());
      }

      function ttMomClass(value){
        const n = safeNum(value);
        if(n > 0.005) return "mom-pos";
        if(n < -0.005) return "mom-neg";
        return "mom-flat";
      }

      function bindTTDownloads(scope){
        if(!scope) return;
        scope.querySelectorAll("[data-plot-download]").forEach((btn)=>{
          btn.onclick = ()=>downloadCanvas(btn.dataset.plotDownload);
        });
      }

      function ensureTTGapCards(){
        const baseCard = document.getElementById("tt-gap-compare-chart")?.closest(".card") || document.getElementById("tt-gap-chart")?.closest(".card");
        if(!baseCard) return null;
        baseCard.classList.add("tt-gap-wide-card");
        baseCard.id = "tt-gap-category-card";
        baseCard.innerHTML = `
          <div class="chart-head"><div><div class="section-label">TT Gap</div><h2>TT销量GAP-品类</h2><div class="muted">左侧看品类销量对比表，右侧看 Shopee 与 TT 当前 vs 上月同期走势</div></div></div>
          <div class="tt-gap-layout">
            <div class="tt-gap-table-panel">
              <div class="tt-gap-table-head"><div class="tt-gap-table-title">品类比货量</div></div>
              <div class="tt-gap-table-wrap" id="tt-gap-category-table"></div>
            </div>
            <div class="tt-gap-chart-panel">
              <div class="tt-gap-chart-stack">
                <div class="tt-gap-chart-block">
                  <div class="tt-gap-chart-head"><div class="tt-gap-chart-title">Shopee-DMS出货趋势</div><button class="analysis-download" data-plot-download="tt-gap-category-shopee-chart">下载图表</button></div>
                  <div id="tt-gap-category-shopee-chart" class="plot"></div>
                </div>
                <div class="tt-gap-chart-block">
                  <div class="tt-gap-chart-head"><div class="tt-gap-chart-title">TT-DMS出货趋势</div><button class="analysis-download" data-plot-download="tt-gap-category-tt-chart">下载图表</button></div>
                  <div id="tt-gap-category-tt-chart" class="plot"></div>
                </div>
              </div>
              <div id="tt-gap-category-note" class="tt-gap-chart-note"></div>
            </div>
          </div>`;
        bindTTDownloads(baseCard);

        let productCard = document.getElementById("tt-gap-product-card");
        if(!productCard){
          productCard = document.createElement("div");
          productCard.className = "card tt-gap-wide-card";
          productCard.id = "tt-gap-product-card";
          baseCard.parentNode.insertBefore(productCard, baseCard.nextSibling);
        }
        productCard.innerHTML = `
          <div class="chart-head"><div><div class="section-label">TT Gap</div><h2>TT销量GAP-产品</h2><div class="muted">左侧看产品销量对比表，右侧看 Shopee 与 TT 当前 vs 上月同期走势</div></div></div>
          <div class="tt-gap-layout">
            <div class="tt-gap-table-panel">
              <div class="tt-gap-table-head"><div class="tt-gap-table-title">产品比货量</div></div>
              <div class="tt-gap-table-wrap" id="tt-gap-product-table"></div>
            </div>
            <div class="tt-gap-chart-panel">
              <div class="tt-gap-chart-stack">
                <div class="tt-gap-chart-block">
                  <div class="tt-gap-chart-head"><div class="tt-gap-chart-title">Shopee-DMS出货趋势</div><button class="analysis-download" data-plot-download="tt-gap-product-shopee-chart">下载图表</button></div>
                  <div id="tt-gap-product-shopee-chart" class="plot"></div>
                </div>
                <div class="tt-gap-chart-block">
                  <div class="tt-gap-chart-head"><div class="tt-gap-chart-title">TT-DMS出货趋势</div><button class="analysis-download" data-plot-download="tt-gap-product-tt-chart">下载图表</button></div>
                  <div id="tt-gap-product-tt-chart" class="plot"></div>
                </div>
              </div>
              <div id="tt-gap-product-note" class="tt-gap-chart-note"></div>
            </div>
          </div>`;
        bindTTDownloads(productCard);
        return { categoryCard: baseCard, productCard };
      }

      function buildTTGapRowsByDim(dim){
        const current = aggregateBy(filterTT(DATA.tt_monthly, "current"), [dim], ["shopee_sales","tt_sales"]);
        const prior = aggregateBy(filterTT(DATA.tt_monthly, "prior"), [dim], ["shopee_sales","tt_sales"]);
        const priorMap = indexBy(prior, dim);
        return current.map((row)=>{
          const key = row[dim];
          if(!key || !ttValidLabel(key)) return null;
          const priorRow = priorMap.get(key) || {};
          const shopee = safeNum(row.shopee_sales);
          const tt = safeNum(row.tt_sales);
          const priorShopee = safeNum(priorRow.shopee_sales);
          const priorTT = safeNum(priorRow.tt_sales);
          const total = shopee + tt;
          const priorTotal = priorShopee + priorTT;
          const gap = shopee - tt;
          const priorGap = priorShopee - priorTT;
          return {
            label:key,
            total,
            totalMom:pctChange(total, priorTotal),
            shopee,
            shopeeMom:pctChange(shopee, priorShopee),
            tt,
            ttMom:pctChange(tt, priorTT),
            gap,
            gapMom:pctChange(gap, priorGap),
            priorShopee,
            priorTT,
            priorTotal,
            priorGap,
          };
        }).filter(Boolean).sort((a,b)=>safeNum(b.total)-safeNum(a.total));
      }

      function ttGapMomHtml(value){
        return `<span class="mom-value ${ttMomClass(value)}">${fmtDelta(value)}</span>`;
      }

      function ttGapValueHtml(value){
        const cls = safeNum(value) >= 0 ? "pos" : "neg";
        return `<span class="tt-gap-value ${cls}">${fmtNum(value)}</span>`;
      }

      function renderTTGapTable(holderId, rows, dimLabel){
        const holder = document.getElementById(holderId);
        if(!holder) return;
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          holder.innerHTML = `<div class="empty-state">当前筛选下暂无${dimLabel}TT销量GAP数据。</div>`;
          return;
        }
        holder.innerHTML = `<table class="tt-gap-table"><thead><tr><th class="tt-rank"></th><th>${dimLabel}</th><th>总计</th><th>% Δ</th><th>SP销量</th><th>% Δ</th><th>TT销量</th><th>% Δ</th><th>GAP值</th><th>% Δ</th></tr></thead><tbody>${list.map((row,idx)=>`<tr><td class="tt-rank">${idx+1}.</td><td>${row.label}</td><td>${fmtNum(row.total)}</td><td>${ttGapMomHtml(row.totalMom)}</td><td>${fmtNum(row.shopee)}</td><td>${ttGapMomHtml(row.shopeeMom)}</td><td>${fmtNum(row.tt)}</td><td>${ttGapMomHtml(row.ttMom)}</td><td>${ttGapValueHtml(row.gap)}</td><td>${ttGapMomHtml(row.gapMom)}</td></tr>`).join("")}</tbody></table>`;
      }

      function drawTTGapTrend(id, rows, field, priorField, currentName, priorName){
        const list = rows.slice(0, Math.max(1, state.topN));
        const el = document.getElementById(id);
        if(!el) return;
        if(!list.length){
          el.innerHTML = '<div class="empty-state">当前筛选下暂无趋势数据。</div>';
          return;
        }
        const labels = list.map((row)=>shortLabel(row.label, 10));
        renderPlot(id,[
          {
            type:"scatter",
            mode:"lines+markers+text",
            name:currentName,
            x:labels,
            y:list.map((row)=>safeNum(row[field])),
            text:list.map((row)=>fmtNum(row[field])),
            textposition:"top center",
            textfont:{color:"#2f6bff",size:11},
            customdata:list.map((row)=>row.label),
            line:{color:"#2f6bff",width:3},
            marker:{size:7,color:"#2f6bff"},
            hovertemplate:`<b>%{customdata}</b><br>${currentName}: %{y:,.0f}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers+text",
            name:priorName,
            x:labels,
            y:list.map((row)=>safeNum(row[priorField])),
            text:list.map((row)=>fmtNum(row[priorField])),
            textposition:"top center",
            textfont:{color:"#9ec6ff",size:11},
            customdata:list.map((row)=>row.label),
            line:{color:"#9ec6ff",width:3},
            marker:{size:7,color:"#9ec6ff"},
            hovertemplate:`<b>%{customdata}</b><br>${priorName}: %{y:,.0f}<extra></extra>`
          }
        ],{
          height:240,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:54,r:16,t:18,b:68},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.16,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function renderTTGapCard(dim){
        const rows = buildTTGapRowsByDim(dim);
        const isCategory = dim === "category";
        const prefix = isCategory ? "category" : "product";
        renderTTGapTable(`tt-gap-${prefix}-table`, rows, isCategory ? "品类" : "产品");
        drawTTGapTrend(`tt-gap-${prefix}-shopee-chart`, rows, "shopee", "priorShopee", "SP销量", "SP销量(上月同期)");
        drawTTGapTrend(`tt-gap-${prefix}-tt-chart`, rows, "tt", "priorTT", "TT销量", "TT销量(上月同期)");
        const note = document.getElementById(`tt-gap-${prefix}-note`);
        if(note){
          const leader = [...rows].sort((a,b)=>safeNum(b.gap)-safeNum(a.gap))[0];
          const lagger = [...rows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];
          if(!rows.length){
            note.innerHTML = '<p>当前筛选下暂无TT销量GAP分析。</p>';
          }else{
            note.innerHTML = `<p>${leader ? `${leader.label} 当前Shopee相对TT领先 ${fmtNum(Math.max(leader.gap,0))}。` : ""} ${lagger ? `${lagger.label} 当前是主要补差点，Shopee较TT${lagger.gap < 0 ? "落后" : "领先"} ${fmtNum(Math.abs(lagger.gap))}。` : ""}</p><p><strong>下一步：</strong>${lagger && lagger.gap < 0 ? `优先围绕 ${lagger.label} 做 TT GAP 补差，补素材、补站内承接和价格钩子。` : "当前缺口相对可控，可继续稳态放量并观察结构变化。"}</p>`;
          }
        }
      }

      window.renderAll = function(reset){
        if(typeof prevRenderAllTTGapDetail === "function") prevRenderAllTTGapDetail(reset);
        ensureTTGapCards();
        renderTTGapCard("category");
        renderTTGapCard("product");
      };

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          ensureTTGapCards();
          if(typeof window.renderAll === "function") window.renderAll(true);
        });
      }else{
        ensureTTGapCards();
        if(typeof window.renderAll === "function") window.renderAll(true);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <style>
    .tt-gap-unified-card{grid-column:1 / -1;display:block !important;width:100%;max-width:none;min-width:0;align-self:stretch;overflow:hidden}
    .chart-grid > .card.tt-gap-unified-card.chart-grid-force-wide{display:block !important;grid-template-columns:none !important;grid-template-areas:none !important;gap:18px !important}
    .chart-grid > .card.tt-gap-unified-card.chart-grid-force-wide > .chart-head{grid-area:auto !important}
    .chart-grid > .card.tt-gap-unified-card.chart-grid-force-wide > .tt-gap-unified-grid{grid-area:auto !important;width:100% !important}
    .tt-gap-unified-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.12fr);width:100%;gap:18px;margin-top:14px;align-items:stretch}
    .tt-gap-unified-left,
    .tt-gap-unified-right{min-width:0;height:100%;width:100%;display:flex;flex-direction:column}
    .tt-gap-unified-left.analysis-block,
    .tt-gap-unified-right.analysis-block{overflow:hidden}
    .tt-gap-unified-stack{display:grid;gap:18px;min-width:0;height:100%;width:100%;flex:1 1 auto}
    .tt-gap-unified-left .tt-gap-unified-stack{grid-template-rows:auto auto}
    .tt-gap-unified-right .tt-gap-unified-stack{grid-template-rows:minmax(0,1fr) minmax(0,1fr) auto}
    .tt-gap-subtable,
    .tt-gap-unified-chart-block{min-width:0}
    .tt-gap-subtable-title{font-size:15px;font-weight:800;color:#244883;line-height:1.4;margin-bottom:10px}
    .tt-gap-subtable + .tt-gap-subtable{margin-top:18px}
    .tt-gap-unified-chart-block{display:flex;flex-direction:column;min-width:0;width:100%}
    .tt-gap-unified-chart-block .plot{flex:1 1 auto;min-height:320px;min-width:0;width:100%}
    .tt-gap-table-wrap{width:100%;max-width:100%;overflow-x:auto;overflow-y:hidden;scrollbar-gutter:stable both-edges;border-radius:16px;border:1px solid rgba(148,163,184,.14);background:#fff}
    .tt-gap-table{width:100%;min-width:0;max-width:none;border-collapse:separate;border-spacing:0;table-layout:fixed}
    .tt-gap-table .mom-value{font-weight:800}
    .tt-gap-table .mom-value.mom-pos{color:#ef4444}
    .tt-gap-table .mom-value.mom-neg{color:#16a34a}
    .tt-gap-table .mom-value.mom-flat{color:#64748b}
    .tt-gap-table .tt-rank{width:34px}
    .tt-gap-table th:nth-child(2),
    .tt-gap-table td:nth-child(2){width:108px;max-width:108px}
    .tt-gap-table th:nth-child(n+3),
    .tt-gap-table td:nth-child(n+3){width:84px}
    .tt-gap-table td:nth-child(2) .tt-label-text,
    .tt-gap-table th:nth-child(2) .tt-label-text{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .tt-gap-mini-note{padding:14px 16px;border-radius:16px;background:#f5f9ff;border:1px solid rgba(46,90,167,.10);font-size:13px;line-height:1.75;color:#334155}
    .tt-gap-mini-note p{margin:0}
    .tt-gap-mini-note p + p{margin-top:6px}
    @media (max-width:1080px){
      .tt-gap-unified-grid{grid-template-columns:1fr}
      .tt-gap-unified-right .tt-gap-unified-stack{grid-template-rows:auto auto auto}
      .tt-gap-unified-chart-block .plot{min-height:280px}
    }
  </style>
  <script>
    (function(){
      const prevRenderAllTTUnified = window.renderAll;
      const TT_LABEL_BLOCKLIST = new Set(["其他","鍏朵粬","未归类","鏈綊绫?","閺堫亜缍婄猾?","闁哄牜浜滅紞濠勭尵?"]);

      function ttUnifiedValidLabel(text){
        return !TT_LABEL_BLOCKLIST.has(String(text || "").trim());
      }

      function ttUnifiedMomClass(value){
        const n = safeNum(value);
        if(n > 0.005) return "mom-pos";
        if(n < -0.005) return "mom-neg";
        return "mom-flat";
      }

      function ttUnifiedEscapeHtml(value){
        return String(value ?? "")
          .replace(/&/g, "&amp;")
          .replace(/"/g, "&quot;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
      }

      function bindUnifiedTTDownloads(scope){
        if(!scope) return;
        scope.querySelectorAll("[data-plot-download]").forEach((btn)=>{
          btn.onclick = ()=>downloadCanvas(btn.dataset.plotDownload);
        });
      }

      function buildTTUnifiedRows(dim){
        const current = aggregateBy(filterTT(DATA.tt_monthly, "current"), [dim], ["shopee_sales","tt_sales"]);
        const prior = aggregateBy(filterTT(DATA.tt_monthly, "prior"), [dim], ["shopee_sales","tt_sales"]);
        const priorMap = indexBy(prior, dim);
        return current.map((row)=>{
          const label = row[dim];
          if(!label || !ttUnifiedValidLabel(label)) return null;
          const priorRow = priorMap.get(label) || {};
          const shopee = safeNum(row.shopee_sales);
          const tt = safeNum(row.tt_sales);
          const priorShopee = safeNum(priorRow.shopee_sales);
          const priorTT = safeNum(priorRow.tt_sales);
          const total = shopee + tt;
          const priorTotal = priorShopee + priorTT;
          const gap = shopee - tt;
          const priorGap = priorShopee - priorTT;
          return {
            label,
            total,
            totalMom:pctChange(total, priorTotal),
            shopee,
            shopeeMom:pctChange(shopee, priorShopee),
            tt,
            ttMom:pctChange(tt, priorTT),
            gap,
            gapMom:pctChange(gap, priorGap),
            priorShopee,
            priorTT
          };
        }).filter(Boolean).sort((a,b)=>safeNum(b.total)-safeNum(a.total));
      }

      function ttUnifiedGapHtml(value){
        const cls = safeNum(value) >= 0 ? "pos" : "neg";
        return `<span class="tt-gap-value ${cls}">${fmtNum(value)}</span>`;
      }

      function ttUnifiedMomHtml(value){
        return `<span class="mom-value ${ttUnifiedMomClass(value)}">${fmtDelta(value)}</span>`;
      }

      function renderTTUnifiedTable(holderId, rows, labelName){
        const holder = document.getElementById(holderId);
        if(!holder) return;
        const list = rows.slice(0, Math.max(1, state.topN));
        if(!list.length){
          holder.innerHTML = `<div class="empty-state">当前筛选下暂无${labelName}TT销量GAP数据。</div>`;
          return;
        }
        holder.innerHTML = `<div class="tt-gap-table-wrap"><table class="tt-gap-table"><thead><tr><th class="tt-rank"></th><th><span class="tt-label-text">${labelName}</span></th><th>SP销量</th><th>% Δ</th><th>TT销量</th><th>% Δ</th><th>GAP值</th><th>% Δ</th></tr></thead><tbody>${list.map((row,idx)=>`<tr><td class="tt-rank">${idx+1}.</td><td title="${ttUnifiedEscapeHtml(row.label)}"><span class="tt-label-text">${row.label}</span></td><td>${fmtNum(row.shopee)}</td><td>${ttUnifiedMomHtml(row.shopeeMom)}</td><td>${fmtNum(row.tt)}</td><td>${ttUnifiedMomHtml(row.ttMom)}</td><td>${ttUnifiedGapHtml(row.gap)}</td><td>${ttUnifiedMomHtml(row.gapMom)}</td></tr>`).join("")}</tbody></table></div>`;
      }

      function ttUnifiedDateLabel(dateText){
        const parts = String(dateText || "").split("-");
        return parts.length === 3 ? `${Number(parts[1])}/${Number(parts[2])}` : String(dateText || "");
      }

      function buildTTUnifiedDailySeries(field){
        const all = Array.isArray(DATA.overview_daily) ? [...DATA.overview_daily] : [];
        const currentRows = all
          .filter((row)=>String(row.date || "") >= String(DATA.meta.current_start || "") && String(row.date || "") <= String(DATA.meta.current_end || ""))
          .sort((a,b)=>String(a.date || "").localeCompare(String(b.date || "")));
        const priorRows = all
          .filter((row)=>String(row.date || "") >= String(DATA.meta.prior_start || "") && String(row.date || "") <= String(DATA.meta.prior_end || ""))
          .sort((a,b)=>String(a.date || "").localeCompare(String(b.date || "")));
        return {
          labels: currentRows.map((row)=>ttUnifiedDateLabel(row.date)),
          currentDates: currentRows.map((row)=>String(row.date || "")),
          currentValues: currentRows.map((row)=>safeNum(row[field])),
          priorDates: priorRows.map((row)=>String(row.date || "")),
          priorValues: priorRows.map((row)=>safeNum(row[field])),
          hasPrior: priorRows.length === currentRows.length && priorRows.length > 0
        };
      }

      function getTTUnifiedTrendHeight(){
        const card = document.getElementById("tt-gap-unified-card");
        const left = card?.querySelector(".tt-gap-unified-left");
        const note = document.getElementById("tt-gap-unified-note");
        const leftHeight = left?.offsetHeight || 0;
        const noteHeight = note?.offsetHeight || 0;
        const available = leftHeight > 0 ? leftHeight - noteHeight - 116 : 660;
        return Math.max(320, Math.floor(available / 2));
      }

      function resizeTTUnifiedPlots(){
        if(!window.Plotly || !Plotly.Plots || typeof Plotly.Plots.resize !== "function") return;
        ["tt-gap-unified-shopee-chart","tt-gap-unified-tt-chart"].forEach((id)=>{
          const el = document.getElementById(id);
          if(!el) return;
          try{
            Plotly.Plots.resize(el);
          }catch(err){
            console.error(err);
          }
        });
      }

      function drawTTUnifiedTrend(id, field, currentName, priorName, chartHeight){
        const holder = document.getElementById(id);
        const series = buildTTUnifiedDailySeries(field);
        if(!holder) return;
        if(!series.labels.length){
          holder.innerHTML = '<div class="empty-state">当前筛选下暂无趋势数据。</div>';
          return;
        }
        const traces = [
          {
            type:"scatter",
            mode:"lines+markers",
            name:currentName,
            x:series.labels,
            y:series.currentValues,
            customdata:series.currentDates,
            line:{color:"#2f6bff",width:3},
            marker:{size:7,color:"#2f6bff"},
            hovertemplate:`<b>%{customdata}</b><br>${currentName}: %{y:,.0f}<extra></extra>`
          }
        ];
        if(series.hasPrior){
          traces.push({
            type:"scatter",
            mode:"lines+markers",
            name:priorName,
            x:series.labels,
            y:series.priorValues,
            customdata:series.priorDates,
            line:{color:"#9ec6ff",width:3},
            marker:{size:7,color:"#9ec6ff"},
            hovertemplate:`<b>%{customdata}</b><br>${priorName}: %{y:,.0f}<extra></extra>`
          });
        }
        renderPlot(id,traces,{
          height:chartHeight || 320,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:54,r:18,t:18,b:54},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:{orientation:"h",x:0,y:1.12,bgcolor:"rgba(255,255,255,.92)",bordercolor:"#dbe3ef",borderwidth:1},
          xaxis:{tickmode:"array",tickvals:series.labels,ticktext:series.labels,tickangle:0,automargin:true,showgrid:false,tickfont:{color:"#64748b"}},
          yaxis:{gridcolor:"#eef3fb",zeroline:false,tickformat:",.0f",tickfont:{color:"#64748b"}}
        },basePlotConfig());
      }

      function mountUnifiedTTCard(){
        const categoryCard = document.getElementById("tt-gap-category-card") || document.getElementById("tt-gap-compare-chart")?.closest(".card") || document.getElementById("tt-gap-chart")?.closest(".card");
        if(!categoryCard) return null;
        categoryCard.className = "card tt-gap-unified-card";
        categoryCard.id = "tt-gap-unified-card";
        categoryCard.innerHTML = `
          <div class="analysis-hidden-legacy">
            <div id="tt-gap-chart"></div>
            <div id="tt-gap-narrative"></div>
            <div id="tt-gap-compare-chart"></div>
            <div id="tt-gap-compare-narrative"></div>
            <div id="tt-gap-delta-chart"></div>
            <div id="tt-gap-delta-narrative"></div>
          </div>
          <div class="chart-head"><div><div class="section-label">TT Gap</div><h2>TT销量GAP</h2><div class="muted">左侧看品类和产品销量对比表，右侧看 Shopee 与 TT 当前 vs 上月同期走势</div></div></div>
          <div class="tt-gap-unified-grid">
            <div class="analysis-block tt-gap-unified-left">
              <div class="tt-gap-unified-stack">
                <div class="tt-gap-subtable">
                  <div class="tt-gap-subtable-title">品类比货量</div>
                  <div id="tt-gap-unified-category-table"></div>
                </div>
                <div class="tt-gap-subtable">
                  <div class="tt-gap-subtable-title">产品比货量</div>
                  <div id="tt-gap-unified-product-table"></div>
                </div>
              </div>
            </div>
            <div class="analysis-block tt-gap-unified-right">
              <div class="tt-gap-unified-stack">
                <div class="tt-gap-unified-chart-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">Shopee-DMS出货趋势</div><div class="analysis-block-note">横轴按日期展示当前期日趋势；如有同期日维度数据则同步叠加上月同期</div></div><button class="analysis-download" data-plot-download="tt-gap-unified-shopee-chart">下载图表</button></div>
                  <div id="tt-gap-unified-shopee-chart" class="plot"></div>
                </div>
                <div class="tt-gap-unified-chart-block">
                  <div class="analysis-block-head"><div><div class="analysis-block-title">TT-DMS出货趋势</div><div class="analysis-block-note">横轴按日期展示当前期日趋势；如有同期日维度数据则同步叠加上月同期</div></div><button class="analysis-download" data-plot-download="tt-gap-unified-tt-chart">下载图表</button></div>
                  <div id="tt-gap-unified-tt-chart" class="plot"></div>
                </div>
                <div id="tt-gap-unified-note" class="tt-gap-mini-note"></div>
              </div>
            </div>
          </div>`;
        bindUnifiedTTDownloads(categoryCard);
        document.getElementById("tt-gap-product-card")?.remove();
        return categoryCard;
      }

      function renderUnifiedTTCard(){
        const card = mountUnifiedTTCard();
        if(!card) return;
        card.classList.remove("chart-grid-single-fill","chart-grid-force-wide");
        const categoryRows = buildTTUnifiedRows("category");
        const productRows = buildTTUnifiedRows("product");
        renderTTUnifiedTable("tt-gap-unified-category-table", categoryRows, "品类");
        renderTTUnifiedTable("tt-gap-unified-product-table", productRows, "产品");
        const note = document.getElementById("tt-gap-unified-note");
        if(note){
          const categoryLagger = [...categoryRows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];
          const productLagger = [...productRows].sort((a,b)=>safeNum(a.gap)-safeNum(b.gap))[0];
          if(!categoryRows.length && !productRows.length){
            note.innerHTML = '<p>当前筛选下暂无TT销量GAP分析。</p>';
          }else{
            note.innerHTML = `<p>${categoryLagger ? `${categoryLagger.label} 当前是品类维度主要补差点，Shopee较TT${categoryLagger.gap < 0 ? "落后" : "领先"} ${fmtNum(Math.abs(categoryLagger.gap))}。` : ""} ${productLagger ? `${productLagger.label} 是产品维度重点关注项。` : ""}</p><p><strong>下一步：</strong>${productLagger && productLagger.gap < 0 ? `优先围绕 ${productLagger.label} 做 TT GAP 补差，补素材、补站内承接和价格钩子。` : "当前缺口相对可控，可继续稳态放量并观察结构变化。"}</p>`;
          }
        }
        const trendHeight = getTTUnifiedTrendHeight();
        drawTTUnifiedTrend("tt-gap-unified-shopee-chart", "shopee_sales", "SP销量", "SP销量(上月同期)", trendHeight);
        drawTTUnifiedTrend("tt-gap-unified-tt-chart", "tt_sales", "TT销量", "TT销量(上月同期)", trendHeight);
        requestAnimationFrame(()=>{
          resizeTTUnifiedPlots();
          setTimeout(resizeTTUnifiedPlots, 80);
        });
      }

      window.renderAll = function(reset){
        if(typeof prevRenderAllTTUnified === "function") prevRenderAllTTUnified(reset);
        renderUnifiedTTCard();
      };

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          renderUnifiedTTCard();
        });
      }else{
        renderUnifiedTTCard();
      }
    })();
  </script>
</body>
""",
        1,
    )
    if plotly_inline_text is not None:
        safe_plotly_inline_text = (
            plotly_inline_text.replace("</script>", "<\\/script>")
            .replace("</SCRIPT>", "<\\/script>")
            .replace("</Script>", "<\\/script>")
        )
        html = html.replace(
            f'<script src="{plotly_src}"></script>',
            f"<script>\n{safe_plotly_inline_text}\n</script>",
            1,
        )
    html = html.replace(
        "</body>",
        """
  <style>
    .module-head-with-filter{
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:16px;
      flex-wrap:wrap;
    }
    .module-filter-tools{
      display:flex;
      align-items:center;
      gap:10px;
      min-width:220px;
      justify-content:flex-end;
    }
    .module-filter-label{
      font-size:12px;
      color:#5f7392;
      white-space:nowrap;
    }
    .module-local-filter{
      height:38px;
      min-width:160px;
      border-radius:12px;
      border:1px solid rgba(46,90,167,.16);
      background:#ffffff;
      color:#244883;
      padding:0 12px;
      font:inherit;
      box-shadow:0 6px 18px rgba(46,90,167,.08);
    }
  </style>
  <script>
    (function(){
      if(typeof state === "object" && state && typeof state.moduleViewFilter === "undefined"){
        state.moduleViewFilter = "";
      }

      function moduleValidLabel(value){
        const text = String(value || "").trim();
        if(!text) return false;
        if(typeof validLabel === "function"){
          try{
            return validLabel(text);
          }catch(err){}
        }
        return !["其他","未归类"].includes(text);
      }

      function moduleEscapeHtml(value){
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;");
      }

      function moduleScopeText(){
        return state.moduleViewFilter ? `${state.moduleViewFilter} 模块` : "全部模块";
      }

      function moduleOptionsSorted(){
        const currentRows = aggregateBy(
          filterRows(DATA.outer_module_daily, "current").filter((row)=>moduleValidLabel(row.module)),
          ["module"],
          ["ad_gmv", "outer_spend", "impressions"]
        );
        return currentRows
          .sort((a,b)=>
            safeNum(b.ad_gmv) - safeNum(a.ad_gmv) ||
            safeNum(b.outer_spend) - safeNum(a.outer_spend) ||
            safeNum(b.impressions) - safeNum(a.impressions) ||
            String(a.module || "").localeCompare(String(b.module || ""), "zh-CN")
          )
          .map((row)=>row.module)
          .filter((value, index, list)=>moduleValidLabel(value) && list.indexOf(value) === index);
      }

      function ensureModuleFilterUi(){
        const card = document.getElementById("module-gmv-cvr-chart")?.closest(".card");
        if(!card) return null;
        const head = card.querySelector(".chart-head");
        if(!head) return null;
        head.classList.add("module-head-with-filter");
        let tools = head.querySelector(".module-filter-tools");
        if(!tools){
          tools = document.createElement("div");
          tools.className = "module-filter-tools";
          tools.innerHTML = '<span class="module-filter-label">模块筛选</span><select id="module-view-filter" class="module-local-filter"></select>';
          head.appendChild(tools);
        }
        const select = tools.querySelector("#module-view-filter");
        if(!select) return null;
        const options = moduleOptionsSorted();
        if(state.moduleViewFilter && !options.includes(state.moduleViewFilter)){
          state.moduleViewFilter = "";
        }
        select.innerHTML = `<option value="">全部模块</option>${options.map((value)=>`<option value="${moduleEscapeHtml(value)}">${moduleEscapeHtml(value)}</option>`).join("")}`;
        select.value = state.moduleViewFilter || "";
        if(!select.dataset.bound){
          select.dataset.bound = "1";
          select.addEventListener("change", (event)=>{
            state.moduleViewFilter = event.target.value || "";
            window.renderAll(false);
          });
        }
        const spendTitle = document.querySelector("#module-gmv-ctr-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-title");
        const spendNote = document.querySelector("#module-gmv-ctr-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        if(spendTitle) spendTitle.textContent = "站外花费柱状 + CTR曲线";
        if(spendNote) spendNote.textContent = "按站外模块口径看点击吸引力与花费强度，支持切换单模块查看日期走势";
        const cpmNote = document.querySelector("#module-cpm-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        const impNote = document.querySelector("#module-impressions-chart")?.closest(".analysis-block")?.querySelector(".analysis-block-note");
        if(cpmNote) cpmNote.textContent = "X轴为日期，可按模块单选查看成本日趋势";
        if(impNote) impNote.textContent = "X轴为日期，可按模块单选查看曝光日趋势";
        return select;
      }

      function buildModuleRowsScoped(){
        const baseRows = typeof buildModuleOuterRowsPure === "function"
          ? buildModuleOuterRowsPure()
          : (typeof buildModuleRows === "function" ? buildModuleRows() : []);
        return (baseRows || []).filter((row)=>moduleValidLabel(row.label) && (!state.moduleViewFilter || row.label === state.moduleViewFilter));
      }

      function buildModuleDailySeries(moduleName){
        const days = Array.from({length:Number(DATA?.meta?.day_count || 0)}, (_, index)=>index + 1);
        const pickRows = (periodType)=>filterRows(DATA.outer_module_daily, periodType).filter((row)=>
          moduleValidLabel(row.module) && (!moduleName || row.module === moduleName)
        );
        const currentMap = indexBy(aggregateBy(pickRows("current"), ["day"], ["outer_spend", "impressions", "clicks"]), "day");
        const priorMap = indexBy(aggregateBy(pickRows("prior"), ["day"], ["outer_spend", "impressions", "clicks"]), "day");
        return {
          axisLabels: chartDayLabels(DATA.meta.current_start, days.length),
          priorAxisLabels: chartDayLabels(DATA.meta.prior_start, days.length),
          currentCpm: days.map((day)=>{
            const row = currentMap.get(day) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            return impressions ? spend / impressions * 1000 : 0;
          }),
          priorCpm: days.map((day)=>{
            const row = priorMap.get(day) || {};
            const spend = safeNum(row.outer_spend);
            const impressions = safeNum(row.impressions);
            return impressions ? spend / impressions * 1000 : 0;
          }),
          currentImpressions: days.map((day)=>safeNum((currentMap.get(day) || {}).impressions)),
          priorImpressions: days.map((day)=>safeNum((priorMap.get(day) || {}).impressions)),
          hasData: days.some((day)=>safeNum((currentMap.get(day) || {}).outer_spend) > 0 || safeNum((currentMap.get(day) || {}).impressions) > 0 || safeNum((priorMap.get(day) || {}).outer_spend) > 0 || safeNum((priorMap.get(day) || {}).impressions) > 0)
        };
      }

      function drawModuleDailyLine(id, axisLabels, priorAxisLabels, currentValues, priorValues, currentName, priorName, currentColor, priorColor, hoverLabel, isPercent){
        const el = document.getElementById(id);
        if(!el) return;
        const hasData = [...currentValues, ...priorValues].some((value)=>safeNum(value) > 0);
        if(!hasData){
          el.innerHTML = '<div class="empty-state">当前筛选下暂无模块日趋势数据。</div>';
          return;
        }
        renderPlot(id,[
          {
            type:"scatter",
            mode:"lines+markers",
            name:currentName,
            x:axisLabels,
            y:currentValues,
            customdata:axisLabels,
            line:{color:currentColor,width:3},
            marker:{size:7,color:currentColor},
            hovertemplate:`<b>${currentName}</b><br>日期：%{customdata}<br>${hoverLabel}：${isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          },
          {
            type:"scatter",
            mode:"lines+markers",
            name:priorName,
            x:axisLabels,
            y:priorValues,
            customdata:priorAxisLabels,
            line:{color:priorColor,width:3},
            marker:{size:7,color:priorColor},
            hovertemplate:`<b>${priorName}</b><br>日期：%{customdata}<br>${hoverLabel}：${isPercent ? "%{y:.1%}" : "%{y:,.0f}"}<extra></extra>`
          }
        ],{
          height:360,
          paper_bgcolor:"rgba(0,0,0,0)",
          plot_bgcolor:"#ffffff",
          margin:{l:76,r:28,t:76,b:100},
          hovermode:"x unified",
          hoverlabel:{bgcolor:"#fff",bordercolor:"#dbe3ef",font:{family:basePlotFont().family,size:12,color:"#1f2937"}},
          font:basePlotFont(),
          legend:chartLegendLayout(),
          xaxis:{
            title:chartAxisTitle("日期"),
            type:"category",
            categoryorder:"array",
            categoryarray:axisLabels,
            tickangle:-38,
            automargin:true,
            showgrid:false,
            linecolor:"#dbe3ef",
            tickfont:{color:"#64748b"}
          },
          yaxis:{
            title:chartAxisTitle(hoverLabel),
            automargin:true,
            gridcolor:"#eef3fb",
            zeroline:false,
            tickformat:isPercent ? ".0%" : ",.0f",
            tickfont:{color:"#64748b"}
          }
        },basePlotConfig());
      }

      function avgValue(list){
        return list.length ? list.reduce((sum, value)=>sum + safeNum(value), 0) / list.length : 0;
      }

      function peakDayLabel(values, labels){
        if(!values.length) return "";
        let peakIndex = 0;
        values.forEach((value, index)=>{
          if(safeNum(value) > safeNum(values[peakIndex])) peakIndex = index;
        });
        return labels[peakIndex] || "";
      }

      function renderModuleSection(){
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"广告GMV",
          priorBarName:"上月同期广告GMV",
          lineName:"CVR",
          priorLineName:"上月同期CVR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"spend",
          priorBarField:"priorSpend",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"站外花费",
          priorBarName:"上月同期站外花费",
          lineName:"CTR",
          priorLineName:"上月同期CTR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        const daily = buildModuleDailySeries(state.moduleViewFilter || "");
        drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "本期CPM", "上月同期CPM", "#2E5AA7", "#88C5FF", "CPM", false);
        drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "本期曝光量", "上月同期曝光量", "#FFA62B", "#F8E6A0", "曝光量", false);

        const scope = moduleScopeText();
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks) > 0).sort((a,b)=>safeNum(b.cvr) - safeNum(a.cvr))[0];
        const topSpend = [...rows].sort((a,b)=>safeNum(b.spend) - safeNum(a.spend))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions) > 0).sort((a,b)=>safeNum(b.ctr) - safeNum(a.ctr))[0];
        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          cvrNarrative.innerHTML = topGmv
            ? `<p>${scope}下，${topGmv.label} 当前广告GMV最高，为 ${fmtMoney(topGmv.gmv)}；${bestCvr ? `${bestCvr.label} 的 CVR 最高，为 ${fmtPct(bestCvr.cvr)}。` : "当前暂无可比较的 CVR 样本。"} 下一步：优先保留高 GMV 且 CVR 不弱的模块预算，避免把量打到点击后承接弱的模块。</p>`
            : '<p>当前筛选下暂无模块广告GMV数据。</p>';
        }
        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          ctrNarrative.innerHTML = topSpend
            ? `<p>${scope}下，${topSpend.label} 当前站外花费最高，为 ${fmtMoney(topSpend.spend)}；${bestCtr ? `${bestCtr.label} 的 CTR 最高，为 ${fmtPct(bestCtr.ctr)}。` : "当前暂无可比较的 CTR 样本。"} 下一步：高花费但 CTR 偏弱的模块优先换素材首屏和卖点钩子，再决定是否继续放量。</p>`
            : '<p>当前筛选下暂无模块花费数据。</p>';
        }
        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels);
            cpmNarrative.innerHTML = `<p>${scope}的 CPM 日均为 ${fmtNum(currentAvg,2)}，较上月同期${priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : " 暂无可比基数"}；${peakLabel || "当前周期"}成本压力最高。 下一步：重点回查高 CPM 日期对应的人群重叠、版位和素材疲劳，避免高成本持续放大。</p>`;
          }else{
            cpmNarrative.innerHTML = '<p>当前筛选下暂无模块 CPM 日趋势数据。</p>';
          }
        }
        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels);
            impressionsNarrative.innerHTML = `<p>${scope}当前累计曝光量为 ${fmtNum(currentTotal)}，较上月同期${priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : " 暂无可比基数"}；${peakLabel || "当前周期"}是曝光峰值日。 下一步：把曝光峰值日与 GMV、点击承接联查，确认高曝光是否真正带来有效成交。</p>`;
          }else{
            impressionsNarrative.innerHTML = '<p>当前筛选下暂无模块曝光日趋势数据。</p>';
          }
        }
      }

      window.renderModuleSectionPatched = renderModuleSection;
      const resetButton = document.getElementById("reset-filter");
      if(resetButton && !resetButton.dataset.moduleResetBound){
        resetButton.dataset.moduleResetBound = "1";
        resetButton.addEventListener("click", ()=>{
          state.moduleViewFilter = "";
        }, true);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      const previousRenderPlotFinalGuard = typeof renderPlot === "function" ? renderPlot : null;
      if(previousRenderPlotFinalGuard){
        renderPlot = function(id, data, layout, config){
          const el = document.getElementById(id);
          if(!el) return null;
          try{
            return previousRenderPlotFinalGuard(id, data, layout, config);
          }catch(err){
            console.error(err);
            return null;
          }
        };
      }
      const previousRenderNarrativesFinalGuard = typeof renderNarratives === "function" ? renderNarratives : null;
      if(previousRenderNarrativesFinalGuard){
        renderNarratives = function(){
          try{
            return previousRenderNarrativesFinalGuard.apply(this, arguments);
          }catch(err){
            console.error(err);
            return null;
          }
        };
      }
      function finalFilterSummaryText(values){
        if(!values || !values.length) return "全部";
        if(values.length <= 2) return values.join(" / ");
        return `${values.slice(0,2).join(" / ")} 等${values.length}项`;
      }

      function finalOrderedOptions(key){
        if(typeof getSortedFilterOptions === "function"){
          try{
            return getSortedFilterOptions(key);
          }catch(err){}
        }
        return [...(DATA.filters[key] || [])];
      }

      function finalSyncNativeSelect(select, values){
        const selected = new Set(values || []);
        Array.from(select.options || []).forEach((option)=>{
          option.selected = selected.has(option.value);
        });
      }

      function finalEnsureMultiSelect(key){
        const select = document.getElementById(`${key}-filter`);
        if(!select || !select.multiple) return null;
        if(typeof ensureMultiSelect === "function"){
          return ensureMultiSelect(key);
        }
        return null;
      }

      function finalRenderMultiSelect(key){
        const select = document.getElementById(`${key}-filter`);
        const wrap = finalEnsureMultiSelect(key);
        if(!select || !wrap) return;
        const orderedOptions = finalOrderedOptions(key);
        if(typeof fillSelect === "function"){
          fillSelect(select, orderedOptions, state.filters[key]);
        }
        const searchInput = wrap.querySelector(".multi-search");
        const keyword = (searchInput?.value || "").trim().toLowerCase();
        const options = orderedOptions.filter((value)=>!keyword || String(value).toLowerCase().includes(keyword));
        const selected = new Set(state.filters[key] || []);
        const triggerText = wrap.querySelector(".multi-trigger-text");
        if(triggerText) triggerText.textContent = finalFilterSummaryText(state.filters[key]);
        const holder = wrap.querySelector(".multi-options");
        if(!holder) return;
        holder.innerHTML = options.map((value)=>{
          const safeValue = String(value).replaceAll('"','&quot;');
          return `<label class="multi-option"><input type="checkbox" value="${safeValue}" ${selected.has(value) ? "checked" : ""}><span class="multi-label">${value}</span><button type="button" class="multi-mini" data-only="${safeValue}">仅限</button></label>`;
        }).join("") || '<div class="empty-state">没有匹配项</div>';
        holder.querySelectorAll('input[type="checkbox"]').forEach((input)=>{
          input.onchange = ()=>{
            const next = new Set(state.filters[key] || []);
            if(input.checked) next.add(input.value);
            else next.delete(input.value);
            state.filters[key] = Array.from(next);
            finalSyncNativeSelect(select, state.filters[key]);
            window.renderAll(false);
            wrap.classList.add("open");
          };
        });
        holder.querySelectorAll("[data-only]").forEach((button)=>{
          button.onclick = ()=>{
            state.filters[key] = [button.dataset.only];
            finalSyncNativeSelect(select, state.filters[key]);
            wrap.classList.remove("open");
            window.renderAll(false);
          };
        });
        if(searchInput && !searchInput.dataset.finalBound){
          searchInput.dataset.finalBound = "1";
          searchInput.addEventListener("input", ()=>finalRenderMultiSelect(key));
        }
      }

      function finalUpdateFilterControls(){
        ["group","category","product","shop"].forEach(finalRenderMultiSelect);
        if(typeof fillSimpleSelect === "function"){
          fillSimpleSelect(document.getElementById("metric-filter"), METRIC_OPTIONS, state.metricFocus);
          fillSimpleSelect(document.getElementById("topn-filter"), TOP_N_OPTIONS.map((value)=>({value, label:String(value)})), state.topN);
        }
        const summary = document.getElementById("filter-summary");
        if(summary){
          summary.textContent = [
            `模块视角：${finalFilterSummaryText(state.filters.group)}`,
            `品类：${finalFilterSummaryText(state.filters.category)}`,
            `产品：${finalFilterSummaryText(state.filters.product)}`,
            `店铺：${finalFilterSummaryText(state.filters.shop)}`,
            `Top N：${state.topN}`
          ].join(" ｜ ");
        }
      }

      function finalRefreshLinkedSections(){
        if(typeof mountSectionCards === "function") mountSectionCards();
        if(typeof renderCategorySection === "function") renderCategorySection();
        if(typeof renderProductSection === "function") renderProductSection();
        if(typeof renderModuleSection === "function") renderModuleSection();
        if(typeof renderTable === "function" && typeof buildDiagnosisRows === "function"){
          renderTable("category", buildDiagnosisRows("category"));
          renderTable("product", buildDiagnosisRows("product"));
        }
        if(typeof buildRoiDiagnosisRows === "function" && typeof renderTable === "function"){
          renderTable("category_roi", buildRoiDiagnosisRows("category_roi"));
          renderTable("product_roi", buildRoiDiagnosisRows("product_roi"));
        }
      }

      function finalBindGlobal(){
        if(window.__dashboardFinalFilterBound) return;
        window.__dashboardFinalFilterBound = true;
        document.addEventListener("click", ()=>document.querySelectorAll(".multi-select.open").forEach((node)=>node.classList.remove("open")));
        const metric = document.getElementById("metric-filter");
        if(metric){
          metric.onchange = (event)=>{
            state.metricFocus = event.target.value;
            window.renderAll(false);
          };
        }
        const topn = document.getElementById("topn-filter");
        if(topn){
          topn.onchange = (event)=>{
            state.topN = Math.min(50, Math.max(1, Number(event.target.value) || 10));
            window.renderAll(false);
          };
        }
        const reset = document.getElementById("reset-filter");
        if(reset){
          reset.onclick = ()=>{
            FILTER_FIELDS.forEach((field)=>{ state.filters[field] = []; });
            document.querySelectorAll(".multi-search").forEach((input)=>{ input.value = ""; });
            state.metricFocus = "gmv";
            state.topN = 10;
            window.renderAll(true);
          };
        }
        document.querySelectorAll("[data-download]").forEach((button)=>{
          button.onclick = ()=>downloadCanvas(button.dataset.download);
        });
      }

      renderMultiSelect = finalRenderMultiSelect;
      updateFilterControls = finalUpdateFilterControls;
      bindGlobal = finalBindGlobal;

      const previousRenderAllFinalSync = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllFinalSync === "function") previousRenderAllFinalSync(reset);
        finalUpdateFilterControls();
        finalRefreshLinkedSections();
      };

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          finalBindGlobal();
          window.renderAll(true);
        });
      }else{
        finalBindGlobal();
        window.renderAll(true);
      }
    })();
  </script>
</body>
""",
        1,
    )
    html = html.replace(
        "</body>",
        """
  <script>
    (function(){
      const previousRenderAllModuleFinal = window.renderAll;
      window.renderAll = function(reset){
        if(typeof previousRenderAllModuleFinal === "function") previousRenderAllModuleFinal(reset);
        if(typeof window.renderModuleSection === "function") window.renderModuleSection();
      };

      function rerenderModuleFinal(){
        if(typeof window.renderModuleSection === "function") window.renderModuleSection();
      }

      if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", function(){
          rerenderModuleFinal();
        });
      }else{
        rerenderModuleFinal();
      }
    })();
  </script>
</body>
""",
        1,
    )
    module_render_start = html.find("function renderModuleSection(){")
    module_render_end = html.find("function renderTTSection(){", module_render_start)
    if module_render_start != -1 and module_render_end != -1:
        html = (
            html[:module_render_start]
            + """function renderModuleSection(){
        if(typeof window.renderModuleSectionPatched === "function"){
          return window.renderModuleSectionPatched();
        }
        return null;
      }

      """
            + html[module_render_end:]
        )
    html = re.sub(
        r'drawModuleDailyLine\("module-cpm-chart",.*?;\n',
        'drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "本期CPM", "上月同期CPM", "#2E5AA7", "#88C5FF", "CPM", false);\n',
        html,
        count=1,
    )
    html = re.sub(
        r'drawModuleDailyLine\("module-impressions-chart",.*?;\n',
        'drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "本期曝光量", "上月同期曝光量", "#FFA62B", "#F8E6A0", "曝光量", false);\n',
        html,
        count=1,
    )
    html = re.sub(
        r"function renderModuleSection\(\)\{\s*ensureModuleFilterUi\(\);[\s\S]*?window\.renderModuleSectionPatched = renderModuleSection;",
        """function renderModuleSection(){
        try{
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"广告GMV",
          priorBarName:"上月同期广告GMV",
          lineName:"CVR",
          priorLineName:"上月同期CVR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"spend",
          priorBarField:"priorSpend",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"站外花费",
          priorBarName:"上月同期站外花费",
          lineName:"CTR",
          priorLineName:"上月同期CTR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        const daily = buildModuleDailySeries(state.moduleViewFilter || "");
        drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "本期CPM", "上月同期CPM", "#2E5AA7", "#88C5FF", "CPM", false);
        drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "本期曝光量", "上月同期曝光量", "#FFA62B", "#F8E6A0", "曝光量", false);

        const scope = moduleScopeText();
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks) > 0).sort((a,b)=>safeNum(b.cvr) - safeNum(a.cvr))[0];
        const topSpend = [...rows].sort((a,b)=>safeNum(b.spend) - safeNum(a.spend))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions) > 0).sort((a,b)=>safeNum(b.ctr) - safeNum(a.ctr))[0];

        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          const cvrText = bestCvr ? `${bestCvr.label} 的CVR最高，为 ${fmtPct(bestCvr.cvr)}。` : "当前暂无可比较的CVR样本。";
          cvrNarrative.innerHTML = topGmv
            ? `<p>${scope}下，${topGmv.label} 的广告GMV最高，为 ${fmtMoney(topGmv.gmv)}；${cvrText} 下一步：优先保留高GMV且CVR稳定的模块预算。</p>`
            : '<p>当前筛选下暂无模块广告GMV数据。</p>';
        }

        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          const ctrText = bestCtr ? `${bestCtr.label} 的CTR最高，为 ${fmtPct(bestCtr.ctr)}。` : "当前暂无可比较的CTR样本。";
          ctrNarrative.innerHTML = topSpend
            ? `<p>${scope}下，${topSpend.label} 的站外花费最高，为 ${fmtMoney(topSpend.spend)}；${ctrText} 下一步：高花费但CTR偏弱的模块优先换素材首屏和卖点钩子。</p>`
            : '<p>当前筛选下暂无模块花费数据。</p>';
        }

        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels) || "当前周期";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "暂无可比基数";
            cpmNarrative.innerHTML = `<p>${scope}的CPM日均为 ${fmtNum(currentAvg,2)}，较上月同期${deltaText}；${peakLabel} 成本压力最高。 下一步：重点回查高CPM日期对应的人群重叠、版位和素材疲劳。</p>`;
          }else{
            cpmNarrative.innerHTML = '<p>当前筛选下暂无模块CPM日趋势数据。</p>';
          }
        }

        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels) || "当前周期";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "暂无可比基数";
            impressionsNarrative.innerHTML = `<p>${scope}当前累计曝光量为 ${fmtNum(currentTotal)}，较上月同期${deltaText}；${peakLabel} 是曝光峰值日。 下一步：把曝光峰值日与GMV、点击承接联查，确认高曝光是否带来有效成交。</p>`;
          }else{
            impressionsNarrative.innerHTML = '<p>当前筛选下暂无模块曝光日趋势数据。</p>';
          }
        }
        }catch(err){
          ["module-gmv-cvr-chart","module-gmv-ctr-chart","module-cpm-chart","module-impressions-chart"].forEach((id)=>{
            const el = document.getElementById(id);
            if(el) el.innerHTML = '<div class="empty-state">Module render error: ' + String(err && err.message ? err.message : err) + '</div>';
          });
          ["module-gmv-cvr-narrative","module-gmv-ctr-narrative","module-cpm-narrative","module-impressions-narrative"].forEach((id)=>{
            const el = document.getElementById(id);
            if(el) el.innerHTML = '<p>Module render error: ' + String(err && err.message ? err.message : err) + '</p>';
          });
        }
      }

      window.renderModuleSectionPatched = renderModuleSection;""",
        html,
        count=1,
    )
    html = re.sub(
        r'drawModuleDailyLine\("module-cpm-chart",.*?;\n',
        'drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "Current CPM", "Prior CPM", "#2E5AA7", "#88C5FF", "CPM", false);\n',
        html,
        count=1,
    )
    html = re.sub(
        r'drawModuleDailyLine\("module-impressions-chart",.*?;\n',
        'drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "Current Impr.", "Prior Impr.", "#FFA62B", "#F8E6A0", "Impressions", false);\n',
        html,
        count=1,
    )
    html = re.sub(
        r"function renderModuleSection\(\)\{\s*ensureModuleFilterUi\(\);[\s\S]*?window\.renderModuleSectionPatched = renderModuleSection;",
        """function renderModuleSection(){
        ensureModuleFilterUi();
        const rows = buildModuleRowsScoped();
        drawBarLineCombo("module-gmv-cvr-chart", rows, {
          barField:"gmv",
          priorBarField:"priorGmv",
          lineField:"cvr",
          priorLineField:"priorCvr",
          barName:"Ad GMV",
          priorBarName:"Prior Ad GMV",
          lineName:"CVR",
          priorLineName:"Prior CVR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        drawBarLineCombo("module-gmv-ctr-chart", rows, {
          barField:"spend",
          priorBarField:"priorSpend",
          lineField:"ctr",
          priorLineField:"priorCtr",
          barName:"Outer Spend",
          priorBarName:"Prior Outer Spend",
          lineName:"CTR",
          priorLineName:"Prior CTR",
          barColor:"#2E5AA7",
          priorBarColor:"#B8DCFF",
          lineColor:"#FFA62B",
          priorLineColor:"#F8E6A0",
          lineIsPercent:true,
        });
        const daily = buildModuleDailySeries(state.moduleViewFilter || "");
        drawModuleDailyLine("module-cpm-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentCpm, daily.priorCpm, "Current CPM", "Prior CPM", "#2E5AA7", "#88C5FF", "CPM", false);
        drawModuleDailyLine("module-impressions-chart", daily.axisLabels, daily.priorAxisLabels, daily.currentImpressions, daily.priorImpressions, "Current Impr.", "Prior Impr.", "#FFA62B", "#F8E6A0", "Impressions", false);

        const scope = moduleScopeText();
        const topGmv = rows[0];
        const bestCvr = [...rows].filter((row)=>safeNum(row.clicks) > 0).sort((a,b)=>safeNum(b.cvr) - safeNum(a.cvr))[0];
        const topSpend = [...rows].sort((a,b)=>safeNum(b.spend) - safeNum(a.spend))[0];
        const bestCtr = [...rows].filter((row)=>safeNum(row.impressions) > 0).sort((a,b)=>safeNum(b.ctr) - safeNum(a.ctr))[0];

        const cvrNarrative = document.getElementById("module-gmv-cvr-narrative");
        if(cvrNarrative){
          const cvrText = bestCvr ? (bestCvr.label + " CVR " + fmtPct(bestCvr.cvr)) : "No CVR sample";
          cvrNarrative.innerHTML = topGmv
            ? ("<p>" + scope + ": " + topGmv.label + " has the highest ad GMV " + fmtMoney(topGmv.gmv) + "; " + cvrText + ". Next: keep budget on modules with both GMV and CVR.</p>")
            : "<p>No module GMV data under current filters.</p>";
        }

        const ctrNarrative = document.getElementById("module-gmv-ctr-narrative");
        if(ctrNarrative){
          const ctrText = bestCtr ? (bestCtr.label + " CTR " + fmtPct(bestCtr.ctr)) : "No CTR sample";
          ctrNarrative.innerHTML = topSpend
            ? ("<p>" + scope + ": " + topSpend.label + " has the highest outer spend " + fmtMoney(topSpend.spend) + "; " + ctrText + ". Next: refresh creative first for high-spend low-CTR modules.</p>")
            : "<p>No module spend data under current filters.</p>";
        }

        const cpmNarrative = document.getElementById("module-cpm-narrative");
        if(cpmNarrative){
          if(daily.hasData){
            const currentAvg = avgValue(daily.currentCpm);
            const priorAvg = avgValue(daily.priorCpm);
            const peakLabel = peakDayLabel(daily.currentCpm, daily.axisLabels) || "current period";
            const deltaText = priorAvg ? fmtDelta(pctChange(currentAvg, priorAvg)) : "no prior base";
            cpmNarrative.innerHTML = "<p>" + scope + " avg CPM " + fmtNum(currentAvg,2) + ", vs prior " + deltaText + "; peak cost day: " + peakLabel + ". Next: check audience overlap, placement and creative fatigue.</p>";
          }else{
            cpmNarrative.innerHTML = "<p>No module CPM daily data under current filters.</p>";
          }
        }

        const impressionsNarrative = document.getElementById("module-impressions-narrative");
        if(impressionsNarrative){
          if(daily.hasData){
            const currentTotal = daily.currentImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const priorTotal = daily.priorImpressions.reduce((sum, value)=>sum + safeNum(value), 0);
            const peakLabel = peakDayLabel(daily.currentImpressions, daily.axisLabels) || "current period";
            const deltaText = priorTotal ? fmtDelta(pctChange(currentTotal, priorTotal)) : "no prior base";
            impressionsNarrative.innerHTML = "<p>" + scope + " total impressions " + fmtNum(currentTotal) + ", vs prior " + deltaText + "; peak exposure day: " + peakLabel + ". Next: link that day back to GMV and click quality.</p>";
          }else{
            impressionsNarrative.innerHTML = "<p>No module impression daily data under current filters.</p>";
          }
        }
      }

      window.renderModuleSectionPatched = renderModuleSection;""",
        html,
        count=1,
    )
    html = html.replace(
        "function drawBarLineCombo(id, rows, options){",
        "window.drawBarLineCombo = function drawBarLineCombo(id, rows, options){",
        1,
    )
    return html


def build_chrome_launcher(target_html: Path) -> str:
    target = str(target_html.resolve())
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        f"set \"TARGET_HTML={target}\"\r\n"
        "set \"CHROME_EXE=C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\"\r\n"
        "if exist \"%CHROME_EXE%\" (\r\n"
        "  start \"\" \"%CHROME_EXE%\" \"%TARGET_HTML%\"\r\n"
        "  exit /b 0\r\n"
        ")\r\n"
        "set \"CHROME_EXE=C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe\"\r\n"
        "if exist \"%CHROME_EXE%\" (\r\n"
        "  start \"\" \"%CHROME_EXE%\" \"%TARGET_HTML%\"\r\n"
        "  exit /b 0\r\n"
        ")\r\n"
        "echo Chrome was not found in the standard install path.\r\n"
        "echo Falling back to the system default browser for this file.\r\n"
        "start \"\" \"%TARGET_HTML%\"\r\n"
    )


def write_chrome_launcher(script_path: Path, target_html: Path) -> Path:
    script_path.parent.mkdir(parents=True, exist_ok=True)
    script_path.write_text(build_chrome_launcher(target_html), encoding="ascii")
    return script_path


def write_desktop_chrome_launcher(target_html: Path) -> Path:
    desktop_path = Path.home() / "Desktop" / "4月Meta Shopee复盘(Chrome打开).cmd"
    return write_chrome_launcher(desktop_path, target_html)


def ensure_plotly_asset(output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / PLOTLY_JS_NAME
    if target.exists() and target.stat().st_size > 1_000_000:
        return target
    with urlopen(PLOTLY_JS_URL) as response:
        target.write_bytes(response.read())
    return target


def save_outputs(payload: dict[str, Any], html_path: str | Path, json_path: str | Path) -> tuple[Path, Path]:
    html_output = Path(html_path)
    json_output = Path(json_path)
    html_output.parent.mkdir(parents=True, exist_ok=True)
    json_output.parent.mkdir(parents=True, exist_ok=True)
    plotly_asset = ensure_plotly_asset(html_output.parent)
    plotly_inline_text = plotly_asset.read_text(encoding="utf-8")
    html_output.write_text(
        render_html(
            payload,
            plotly_src=f"./{plotly_asset.name}",
            plotly_inline_text=plotly_inline_text,
        ),
        encoding="utf-8",
    )
    json_output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_chrome_launcher(
        html_output.parent / f"open_{html_output.stem}_in_chrome.cmd",
        html_output,
    )
    write_chrome_launcher(
        ROOT / "双击用Chrome打开4月复盘.cmd",
        html_output,
    )
    write_chrome_launcher(
        html_output.parent / "双击用Chrome打开4月复盘.cmd",
        html_output,
    )
    write_desktop_chrome_launcher(html_output)
    return html_output, json_output


def main() -> None:
    args = parse_args()
    config = load_source_config(args.source_config)
    payload = build_payload(config=config, workbook_override=args.workbook)
    save_outputs(payload=payload, html_path=args.output_html, json_path=args.output_json)


if __name__ == "__main__":
    main()
