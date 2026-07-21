from __future__ import annotations

from typing import Iterable

import numpy as np
import pandas as pd

from skills.clean_data import ALL_METRICS, CORE_METRICS, ensure_columns


def safe_divide(
    numerator: pd.Series | float | int | None,
    denominator: pd.Series | float | int | None,
) -> pd.Series | float:
    if isinstance(numerator, pd.Series) or isinstance(denominator, pd.Series):
        if not isinstance(numerator, pd.Series):
            numerator = pd.Series(numerator, index=denominator.index if isinstance(denominator, pd.Series) else None)
        if not isinstance(denominator, pd.Series):
            denominator = pd.Series(denominator, index=numerator.index)

        result = numerator.astype(float).divide(denominator.astype(float).replace({0: np.nan}))
        return result.replace([np.inf, -np.inf], np.nan)

    if denominator in (0, None) or pd.isna(denominator):
        return np.nan
    if numerator is None or pd.isna(numerator):
        return np.nan
    return float(numerator) / float(denominator)


def _sum_metric(df: pd.DataFrame, column: str) -> float:
    if column not in df.columns:
        return np.nan
    return df[column].sum(min_count=1)


def summarize_raw_metrics(
    df: pd.DataFrame,
    raw_metrics: Iterable[str] | None = None,
) -> pd.DataFrame:
    raw_metrics = list(raw_metrics or CORE_METRICS)
    prepared = ensure_columns(df, raw_metrics)
    summary = {metric: _sum_metric(prepared, metric) for metric in raw_metrics}
    return pd.DataFrame([summary])


def compute_metrics(
    df: pd.DataFrame,
    raw_metrics: Iterable[str] | None = None,
) -> pd.DataFrame:
    raw_metrics = list(raw_metrics or CORE_METRICS)
    prepared = summarize_raw_metrics(df, raw_metrics=raw_metrics)
    prepared = ensure_columns(prepared, ALL_METRICS)

    prepared["ctr"] = safe_divide(prepared["clicks"], prepared["impressions"])
    prepared["cpc"] = safe_divide(prepared["spend"], prepared["clicks"])
    prepared["cpm"] = safe_divide(prepared["spend"], prepared["impressions"]) * 1000
    prepared["roi"] = safe_divide(prepared["gmv"], prepared["spend"])
    prepared["aov"] = safe_divide(prepared["gmv"], prepared["orders"])
    prepared["cvr"] = safe_divide(prepared["orders"], prepared["visitors"])

    return prepared[ALL_METRICS]


def compute_metrics_by_dimension(
    df: pd.DataFrame,
    dimension: str,
    raw_metrics: Iterable[str] | None = None,
) -> pd.DataFrame:
    raw_metrics = list(raw_metrics or CORE_METRICS)
    if dimension not in df.columns:
        columns = [dimension] + ALL_METRICS
        return pd.DataFrame(columns=columns)

    grouped = (
        ensure_columns(df, raw_metrics)
        .groupby(dimension, dropna=False)[raw_metrics]
        .sum(min_count=1)
        .reset_index()
    )

    grouped["ctr"] = safe_divide(grouped["clicks"], grouped["impressions"])
    grouped["cpc"] = safe_divide(grouped["spend"], grouped["clicks"])
    grouped["cpm"] = safe_divide(grouped["spend"], grouped["impressions"]) * 1000
    grouped["roi"] = safe_divide(grouped["gmv"], grouped["spend"])
    grouped["aov"] = safe_divide(grouped["gmv"], grouped["orders"])
    grouped["cvr"] = safe_divide(grouped["orders"], grouped["visitors"])

    return grouped[[dimension] + ALL_METRICS]
