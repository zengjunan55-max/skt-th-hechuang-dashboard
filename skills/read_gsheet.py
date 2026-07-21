from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:  # pragma: no cover - optional at test time
    gspread = None
    Credentials = None


GOOGLE_SHEET_KEY_RE = re.compile(r"(?:/spreadsheets/d/|^)(?P<key>[-\w]{20,})(?:/|$)")
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def extract_sheet_key(source: str) -> str:
    match = GOOGLE_SHEET_KEY_RE.search(str(source).strip())
    if not match:
        raise ValueError(f"Invalid Google Sheet source: {source}")
    return match.group("key")


def build_gspread_client(
    credentials_path: str | Path,
    scopes: list[str] | None = None,
) -> Any:
    if gspread is None or Credentials is None:
        raise ImportError("gspread and google-auth are required to read Google Sheets.")

    credentials = Credentials.from_service_account_file(
        str(credentials_path),
        scopes=scopes or GOOGLE_SCOPES,
    )
    return gspread.authorize(credentials)


def read_gsheet(
    source: str,
    worksheet: str | int = 0,
    credentials_path: str | Path | None = None,
    client: Any | None = None,
) -> pd.DataFrame:
    """
    Read a single Google Sheet tab into a pandas DataFrame.

    Pass either `credentials_path` or an already-authenticated `client`.
    """
    if client is None:
        if credentials_path is None:
            raise ValueError("credentials_path is required when client is not provided.")
        client = build_gspread_client(credentials_path)

    spreadsheet = client.open_by_key(extract_sheet_key(source))

    if isinstance(worksheet, int):
        sheet = spreadsheet.get_worksheet(worksheet)
        if sheet is None:
            raise ValueError(f"Worksheet index not found: {worksheet}")
    else:
        sheet = spreadsheet.worksheet(worksheet)

    records = sheet.get_all_records()
    return pd.DataFrame(records)


def is_google_sheet_source(source: Any) -> bool:
    if source is None:
        return False
    text = str(source).strip()
    return "docs.google.com/spreadsheets" in text or bool(GOOGLE_SHEET_KEY_RE.fullmatch(text))


def read_tabular_data(
    source: str | Path | pd.DataFrame,
    worksheet: str | int = 0,
    credentials_path: str | Path | None = None,
    client: Any | None = None,
    **read_kwargs: Any,
) -> pd.DataFrame:
    if isinstance(source, pd.DataFrame):
        return source.copy()

    if source is None:
        raise ValueError("source is required.")

    if is_google_sheet_source(source):
        return read_gsheet(
            source=str(source),
            worksheet=worksheet,
            credentials_path=credentials_path,
            client=client,
        )

    path = Path(source)
    if not path.exists():
        raise FileNotFoundError(f"Data file not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, **read_kwargs)
    if suffix in {".xlsx", ".xls"}:
        excel_kwargs = {"sheet_name": worksheet}
        excel_kwargs.update(read_kwargs)
        return pd.read_excel(path, **excel_kwargs)
    if suffix == ".parquet":
        return pd.read_parquet(path, **read_kwargs)

    raise ValueError(f"Unsupported file type: {suffix}")
