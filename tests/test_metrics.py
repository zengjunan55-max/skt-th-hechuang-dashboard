from __future__ import annotations

import math
import sys
from pathlib import Path
from unittest.mock import Mock

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from skills.clean_data import CORE_METRICS, get_missing_columns, prepare_analysis_data
from skills.compute_metrics import compute_metrics, compute_metrics_by_dimension, safe_divide
from skills.read_gsheet import extract_sheet_key, read_gsheet


def test_extract_sheet_key_supports_url_and_bare_key() -> None:
    key = "1AbCdEfGhIjKlMnOpQrStUvWxYz123456"
    assert extract_sheet_key(key) == key
    assert (
        extract_sheet_key(f"https://docs.google.com/spreadsheets/d/{key}/edit#gid=0")
        == key
    )


def test_read_gsheet_returns_dataframe_from_named_tab() -> None:
    worksheet = Mock()
    worksheet.get_all_records.return_value = [{"Spend": "100", "Clicks": "10"}]

    spreadsheet = Mock()
    spreadsheet.worksheet.return_value = worksheet

    client = Mock()
    client.open_by_key.return_value = spreadsheet

    df = read_gsheet(
        source="https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz123456/edit",
        worksheet="Sheet1",
        client=client,
    )

    assert list(df.columns) == ["Spend", "Clicks"]
    assert df.iloc[0]["Spend"] == "100"
    client.open_by_key.assert_called_once()
    spreadsheet.worksheet.assert_called_once_with("Sheet1")


def test_prepare_analysis_data_normalizes_columns_and_types() -> None:
    raw = pd.DataFrame(
        [
            {
                " Spend ": "100",
                "Impressions": "1000",
                "Clicks": "50",
                "Landing Page Views": "25",
                "Purchases": "5",
                "Revenue": "300",
            }
        ]
    )

    prepared = prepare_analysis_data(raw, required_columns=CORE_METRICS)

    assert list(prepared.columns) == [
        "spend",
        "impressions",
        "clicks",
        "visitors",
        "orders",
        "gmv",
    ]
    assert prepared.iloc[0]["spend"] == 100
    assert prepared.iloc[0]["visitors"] == 25
    assert prepared.iloc[0]["orders"] == 5
    assert prepared.iloc[0]["gmv"] == 300


def test_prepare_analysis_data_adds_missing_core_columns() -> None:
    raw = pd.DataFrame([{"Spend": 100, "Clicks": 10}])
    prepared = prepare_analysis_data(raw, required_columns=CORE_METRICS)

    assert get_missing_columns(prepared, CORE_METRICS) == []
    assert math.isnan(prepared.iloc[0]["impressions"])
    assert math.isnan(prepared.iloc[0]["gmv"])


def test_safe_divide_returns_nan_on_zero_denominator() -> None:
    result = safe_divide(10, 0)
    assert math.isnan(result)


def test_compute_metrics_builds_all_required_metrics() -> None:
    df = pd.DataFrame(
        [
            {
                "spend": 100,
                "impressions": 1000,
                "clicks": 100,
                "visitors": 50,
                "orders": 10,
                "gmv": 300,
            }
        ]
    )

    result = compute_metrics(df)
    row = result.iloc[0]

    assert row["spend"] == 100
    assert row["impressions"] == 1000
    assert row["clicks"] == 100
    assert row["ctr"] == 0.1
    assert row["cpc"] == 1
    assert row["cpm"] == 100
    assert row["visitors"] == 50
    assert row["orders"] == 10
    assert row["gmv"] == 300
    assert row["roi"] == 3
    assert row["aov"] == 30
    assert row["cvr"] == 0.2


def test_compute_metrics_handles_missing_columns_and_zero_values() -> None:
    df = pd.DataFrame([{"spend": 100, "clicks": 0}])
    result = compute_metrics(df).iloc[0]

    assert result["spend"] == 100
    assert result["clicks"] == 0
    assert math.isnan(result["impressions"])
    assert math.isnan(result["cpc"])
    assert math.isnan(result["ctr"])
    assert math.isnan(result["roi"])


def test_compute_metrics_by_dimension_groups_rows() -> None:
    df = pd.DataFrame(
        [
            {
                "module": "A",
                "spend": 100,
                "impressions": 1000,
                "clicks": 100,
                "visitors": 60,
                "orders": 6,
                "gmv": 300,
            },
            {
                "module": "A",
                "spend": 50,
                "impressions": 500,
                "clicks": 50,
                "visitors": 20,
                "orders": 2,
                "gmv": 120,
            },
        ]
    )

    result = compute_metrics_by_dimension(df, "module")
    row = result.iloc[0]

    assert row["module"] == "A"
    assert row["spend"] == 150
    assert row["gmv"] == 420
    assert row["roi"] == 2.8
    assert row["cvr"] == 0.1
