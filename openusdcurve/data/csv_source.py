"""User-supplied CSV quote ingestion.

CSV schema (header row required, comma-separated, UTF-8). Column names are case-insensitive.

Required columns
-----------------
- ``valuation_date``   ISO date ``YYYY-MM-DD``.
- ``instrument_type``  One of: deposit, future, libor_swap, ois, fixing, sofr_index,
                       treasury_par (matches ``openusdcurve.data.base.InstrumentType`` values).
- ``instrument_id``    Free-text identifier, e.g. ``DEPO_3M``, ``SOFR_FUT_H24``.
- ``maturity_date``    ISO date ``YYYY-MM-DD``.
- ``quote``            Numeric. Rates/par-rates are DECIMALS (0.0525 == 5.25%), never percent.
                       Futures use price (e.g. 95.25).
- ``quote_type``       One of: rate, price, par_rate, index (matches ``QuoteType``).

Optional columns
----------------
- ``start_date``    ISO date; accrual start (defaults to ``valuation_date`` if omitted).
- ``bid``, ``ask``  Numeric, same units as ``quote``.
- ``source_url``    Free text.
- ``delayed``       ``true``/``false`` (default false).
- ``provisional``   ``true``/``false`` (default false).
- ``observed_at``   ISO 8601 datetime, e.g. ``2026-07-20T14:30:00``.

Rows that fail to parse a required field raise ``ValueError`` naming the row and column so bad
manual data is caught at load time rather than silently propagated into a bootstrap. An example
file following this schema ships at ``data/sample/example_quotes.csv``.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["CSVSource", "REQUIRED_COLUMNS", "OPTIONAL_COLUMNS"]

REQUIRED_COLUMNS = (
    "valuation_date",
    "instrument_type",
    "instrument_id",
    "maturity_date",
    "quote",
    "quote_type",
)
OPTIONAL_COLUMNS = (
    "start_date",
    "bid",
    "ask",
    "source_url",
    "delayed",
    "provisional",
    "observed_at",
)


def _parse_date(value: object, *, row: int, column: str) -> date:
    if pd.isna(value):
        raise ValueError(f"row {row}: required column '{column}' is empty")
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    text = str(value).strip()
    try:
        return date.fromisoformat(text[:10])
    except ValueError as exc:
        raise ValueError(f"row {row}: could not parse date '{value}' in column '{column}'") from exc


def _parse_bool(value: object) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def _parse_optional_float(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    return float(value)


@dataclass
class CSVSource:
    """Reads a schema-conforming CSV of quotes into normalized :class:`MarketQuote` objects.

    All quotes are tagged ``license=MANUAL_USER_SUPPLIED`` regardless of any license info in the
    file, since provenance cannot be independently verified for arbitrary user uploads.
    """

    path: str | Path
    name: str = "csv"
    license: LicenseTag = LicenseTag.MANUAL_USER_SUPPLIED

    def get_quotes(self, valuation_date: date | None = None) -> list[MarketQuote]:
        """Load all quotes from the CSV. If ``valuation_date`` is given, filter to that date."""
        df = pd.read_csv(self.path)
        df.columns = [c.strip().lower() for c in df.columns]
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                f"{self.path}: missing required column(s) {missing}; "
                f"required={REQUIRED_COLUMNS}"
            )

        quotes: list[MarketQuote] = []
        for i, row in df.iterrows():
            row_num = int(i) + 2  # +1 header, +1 to make it 1-indexed for humans
            vdate = _parse_date(row["valuation_date"], row=row_num, column="valuation_date")
            mdate = _parse_date(row["maturity_date"], row=row_num, column="maturity_date")

            try:
                instrument_type = InstrumentType(str(row["instrument_type"]).strip().lower())
            except ValueError as exc:
                raise ValueError(
                    f"row {row_num}: invalid instrument_type '{row['instrument_type']}'"
                ) from exc
            try:
                quote_type = QuoteType(str(row["quote_type"]).strip().lower())
            except ValueError as exc:
                raise ValueError(
                    f"row {row_num}: invalid quote_type '{row['quote_type']}'"
                ) from exc

            if pd.isna(row["quote"]):
                raise ValueError(f"row {row_num}: 'quote' is empty")

            start_date = None
            if "start_date" in df.columns and not pd.isna(row.get("start_date")):
                start_date = _parse_date(row["start_date"], row=row_num, column="start_date")

            observed_at = None
            if "observed_at" in df.columns and not pd.isna(row.get("observed_at")):
                observed_at = datetime.fromisoformat(str(row["observed_at"]).strip())

            quote = MarketQuote(
                valuation_date=vdate,
                instrument_type=instrument_type,
                instrument_id=str(row["instrument_id"]).strip(),
                maturity_date=mdate,
                quote=float(row["quote"]),
                quote_type=quote_type,
                source=self.name,
                license=self.license,
                observed_at=observed_at,
                start_date=start_date,
                bid=_parse_optional_float(row.get("bid")) if "bid" in df.columns else None,
                ask=_parse_optional_float(row.get("ask")) if "ask" in df.columns else None,
                source_url=(
                    str(row["source_url"]).strip()
                    if "source_url" in df.columns and not pd.isna(row.get("source_url"))
                    else None
                ),
                delayed=_parse_bool(row.get("delayed")) if "delayed" in df.columns else False,
                provisional=(
                    _parse_bool(row.get("provisional")) if "provisional" in df.columns else False
                ),
            )
            quotes.append(quote)

        if valuation_date is not None:
            quotes = [q for q in quotes if q.valuation_date == valuation_date]
        return quotes
