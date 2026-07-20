"""US Treasury daily par yield curve.

Public CSV endpoint (no key required), pattern:

    https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/{year}/all?type=daily_treasury_yield_curve&field_tdr_date_value={year}&page&_format=csv

Columns vary slightly by year (Treasury has added/retired tenors over time, e.g. 4 Mo added
2018-10, 20 Yr reintroduced 2020-05), but always include a ``Date`` column plus one column per
published tenor, values in percent (e.g. ``4.21`` == 4.21%). Recognized tenor column headers:
``1 Mo``, ``1.5 Month``, ``2 Mo``, ``3 Mo``, ``4 Mo``, ``6 Mo``, ``1 Yr``, ``2 Yr``, ``3 Yr``,
``5 Yr``, ``7 Yr``, ``10 Yr``, ``20 Yr``, ``30 Yr``.

TODO: confirm the exact current query-string contract with a live pull; Treasury has changed
this endpoint's parameters before (an older XML endpoint also exists). The parsing helpers below
operate on already-downloaded CSV text, so a URL change only requires updating ``_url_for_year``.

Design constraints: import must not perform network I/O; the HTTP client is injectable so
parsing can be tested fully offline.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["TreasurySource"]

_BASE_URL = "https://home.treasury.gov"
_PATH_TEMPLATE = (
    "/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/{year}/all"
)

_TENOR_MONTHS = {
    "1 mo": 1,
    "1.5 month": 1.5,
    "2 mo": 2,
    "3 mo": 3,
    "4 mo": 4,
    "6 mo": 6,
    "1 yr": 12,
    "2 yr": 24,
    "3 yr": 36,
    "5 yr": 60,
    "7 yr": 84,
    "10 yr": 120,
    "20 yr": 240,
    "30 yr": 360,
}


def _url_for_year(year: int) -> tuple[str, dict]:
    path = _PATH_TEMPLATE.format(year=year)
    params = {"type": "daily_treasury_yield_curve", "field_tdr_date_value": year, "_format": "csv"}
    return path, params


def parse_treasury_csv(text: str) -> pd.DataFrame:
    """Parse a Treasury daily par yield curve CSV into a tidy long-format DataFrame with columns
    ``date``, ``tenor_label``, ``tenor_months``, ``rate`` (decimal, i.e. percent / 100)."""
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip() for c in df.columns]
    date_col = df.columns[0]
    df[date_col] = pd.to_datetime(df[date_col]).dt.date

    tenor_cols = [c for c in df.columns if c != date_col]
    records = []
    for _, row in df.iterrows():
        for col in tenor_cols:
            key = re.sub(r"\s+", " ", col.strip().lower())
            if key not in _TENOR_MONTHS or pd.isna(row[col]):
                continue
            records.append(
                {
                    "date": row[date_col],
                    "tenor_label": col.strip(),
                    "tenor_months": _TENOR_MONTHS[key],
                    "rate": float(row[col]) / 100.0,
                }
            )
    return pd.DataFrame(records)


def _add_months(d: date, months: float) -> date:
    import calendar

    total = d.month - 1 + int(round(months))
    year = d.year + total // 12
    month = total % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


@dataclass
class TreasurySource:
    """US Treasury daily par yield curve. Returns InstrumentType.TREASURY_PAR MarketQuotes."""

    base_url: str = _BASE_URL
    client: object | None = None
    name: str = "us-treasury"
    license: LicenseTag = LicenseTag.PUBLIC_DOMAIN

    _owns_client: bool = field(default=False, init=False, repr=False)

    def _get_client(self):
        if self.client is None:
            import httpx

            self.client = httpx.Client(base_url=self.base_url, timeout=30.0)
            self._owns_client = True
        return self.client

    def _download_year_csv(self, year: int) -> str:
        path, params = _url_for_year(year)
        client = self._get_client()
        response = client.get(path, params=params)
        response.raise_for_status()
        return response.text

    def get_par_curve(self, valuation_date: date) -> pd.DataFrame:
        """Par yield curve rows for the single ``valuation_date`` (empty if not a published
        business day in the downloaded year's data)."""
        text = self._download_year_csv(valuation_date.year)
        df = parse_treasury_csv(text)
        return df[df["date"] == valuation_date].reset_index(drop=True)

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        df = self.get_par_curve(valuation_date)
        out = []
        for _, row in df.iterrows():
            maturity = _add_months(valuation_date, row["tenor_months"])
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.TREASURY_PAR,
                    instrument_id=f"UST_{row['tenor_label'].replace(' ', '')}",
                    maturity_date=maturity,
                    quote=float(row["rate"]),
                    quote_type=QuoteType.RATE,
                    source=self.name,
                    license=self.license,
                    source_url=f"{self.base_url}{_url_for_year(valuation_date.year)[0]}",
                )
            )
        return out
