from __future__ import annotations

import re
from typing import Iterable

import pandas as pd


CANONICAL_COLUMN_ALIASES = {
    "spend": "spend",
    "amount_spent": "spend",
    "impressions": "impressions",
    "clicks": "clicks",
    "ctr": "ctr",
    "cpc": "cpc",
    "cpm": "cpm",
    "visitors": "visitors",
    "sessions": "visitors",
    "landing_page_views": "visitors",
    "orders": "orders",
    "purchases": "orders",
    "gmv": "gmv",
    "revenue": "gmv",
    "sales": "gmv",
    "roi": "roi",
    "roas": "roi",
    "aov": "aov",
    "average_order_value": "aov",
    "cvr": "cvr",
    "conversion_rate": "cvr",
}

CORE_METRICS = [
    "spend",
    "impressions",
    "clicks",
    "visitors",
    "orders",
    "gmv",
]

ALL_METRICS = [
    "spend",
    "impressions",
    "clicks",
    "ctr",
    "cpc",
    "cpm",
    "visitors",
    "orders",
    "gmv",
    "roi",
    "aov",
    "cvr",
]


def standardize_column_name(column_name: object) -> str:
    text = str(column_name).strip().lower()
    text = re.sub(r"[%()/]+", " ", text)
    text = re.sub(r"[^0-9a-zA-Z_]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return CANONICAL_COLUMN_ALIASES.get(text, text)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized = df.copy()
    normalized.columns = [standardize_column_name(column) for column in normalized.columns]
    return normalized


def replace_blank_strings(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    object_columns = cleaned.select_dtypes(include=["object", "string"]).columns
    for column in object_columns:
        cleaned[column] = cleaned[column].replace(r"^\s*$", pd.NA, regex=True)
    return cleaned


def coerce_numeric_columns(
    df: pd.DataFrame,
    numeric_columns: Iterable[str],
) -> pd.DataFrame:
    cleaned = df.copy()
    for column in numeric_columns:
        if column in cleaned.columns:
            cleaned[column] = pd.to_numeric(cleaned[column], errors="coerce")
    return cleaned


def ensure_columns(
    df: pd.DataFrame,
    required_columns: Iterable[str],
    fill_value: object = pd.NA,
) -> pd.DataFrame:
    prepared = df.copy()
    for column in required_columns:
        if column not in prepared.columns:
            prepared[column] = fill_value
    return prepared


def get_missing_columns(
    df: pd.DataFrame,
    required_columns: Iterable[str],
) -> list[str]:
    existing = set(df.columns)
    return [column for column in required_columns if column not in existing]


def normalize_data_types(
    df: pd.DataFrame,
    numeric_columns: Iterable[str] | None = None,
) -> pd.DataFrame:
    numeric_columns = list(numeric_columns or CORE_METRICS)
    cleaned = replace_blank_strings(df)
    cleaned = ensure_columns(cleaned, numeric_columns)
    cleaned = coerce_numeric_columns(cleaned, numeric_columns)
    return cleaned


def prepare_analysis_data(
    df: pd.DataFrame,
    required_columns: Iterable[str] | None = None,
) -> pd.DataFrame:
    prepared = normalize_columns(df)
    prepared = normalize_data_types(prepared, numeric_columns=required_columns or CORE_METRICS)
    return prepared
