"""FRED (Federal Reserve Economic Data) backup source.

Uses the public ``fredgraph.csv`` download endpoint, which requires no API key:

    https://fred.stlouisfed.org/graph/fredgraph.csv?id=<SERIES_ID>&cosd=<start>&coed=<end>

Series used (FRED series IDs):

- ``SOFR``          Secured Overnight Financing Rate (daily), backup to NY Fed.
- ``USD1MTD156N``   1-Month USD LIBOR (historical, series discontinued 2023-06 but archived).
- ``USD3MTD156N``   3-Month USD LIBOR (historical).
- ``USD6MTD156N``   6-Month USD LIBOR (historical).
- ``USD12MTD156N``  12-Month USD LIBOR (historical).
- ``DSWP2`` / ``DSWP5`` / ``DSWP10`` / ``DSWP30``  2Y/5Y/10Y/30Y USD interest rate swap rates
  (historical, series discontinued 2016 but archived; used for OS-2 Lehman-era reconstruction).

TODO: FRED occasionally renames/retires series; if a series ID 404s, check
https://fred.stlouisfed.org for the current ID before assuming the data is gone.

Design constraints: importing this module must never perform network I/O; the HTTP
transport is injectable, and parsing is exercised in tests against a raw CSV string / a
committed sample file with no network access.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["FREDSource", "SERIES_IDS"]

_BASE_URL = "https://fred.stlouisfed.org"
_GRAPH_PATH = "/graph/fredgraph.csv"

SERIES_IDS = {
    "sofr": "SOFR",
    "libor_1m": "USD1MTD156N",
    "libor_3m": "USD3MTD156N",
    "libor_6m": "USD6MTD156N",
    "libor_12m": "USD12MTD156N",
    "swap_2y": "DSWP2",
    "swap_5y": "DSWP5",
    "swap_10y": "DSWP10",
    "swap_30y": "DSWP30",
}

_TENOR_TO_LIBOR_KEY = {1: "libor_1m", 3: "libor_3m", 6: "libor_6m", 12: "libor_12m"}
_TENOR_TO_SWAP_KEY = {2: "swap_2y", 5: "swap_5y", 10: "swap_10y", 30: "swap_30y"}


def parse_fredgraph_csv(text: str, series_id: str) -> pd.DataFrame:
    """Parse fredgraph.csv text (two columns: DATE, <series_id>) into a tidy DataFrame with
    columns ``date`` (datetime.date) and ``value`` (float, NaN for the ``.`` missing marker).
    Values that look like percent (e.g. LIBOR/swap series) are left as-is here — FRED already
    reports these series in percent, so callers must divide by 100 to get decimals, which
    ``FREDSource`` does when building MarketQuotes."""
    df = pd.read_csv(io.StringIO(text), na_values=["."])
    df.columns = [c.strip() for c in df.columns]
    date_col = df.columns[0]
    value_col = series_id if series_id in df.columns else df.columns[1]
    out = pd.DataFrame(
        {
            "date": pd.to_datetime(df[date_col]).dt.date,
            "value": pd.to_numeric(df[value_col], errors="coerce"),
        }
    )
    return out.dropna(subset=["value"]).reset_index(drop=True)


@dataclass
class FREDSource:
    """Backup daily SOFR + historical USD swap rates + historical LIBOR via FRED CSV endpoint.

    ``client`` is injectable (an ``httpx.Client``) so tests can supply a mock transport; no
    network I/O happens at import or construction time.
    """

    base_url: str = _BASE_URL
    client: object | None = None
    name: str = "fred"
    license: LicenseTag = LicenseTag.PUBLIC_DOMAIN

    _owns_client: bool = field(default=False, init=False, repr=False)

    def _get_client(self):
        if self.client is None:
            import httpx

            self.client = httpx.Client(base_url=self.base_url, timeout=30.0)
            self._owns_client = True
        return self.client

    def _download_csv_text(self, series_id: str, start: date, end: date) -> str:
        client = self._get_client()
        response = client.get(
            _GRAPH_PATH,
            params={"id": series_id, "cosd": start.isoformat(), "coed": end.isoformat()},
        )
        response.raise_for_status()
        return response.text

    def get_series(self, series_key: str, start: date, end: date) -> pd.DataFrame:
        """``series_key`` is a key of ``SERIES_IDS`` (e.g. 'sofr', 'libor_3m', 'swap_10y')."""
        series_id = SERIES_IDS[series_key]
        text = self._download_csv_text(series_id, start, end)
        return parse_fredgraph_csv(text, series_id)

    def get_sofr(self, start: date, end: date) -> pd.DataFrame:
        return self.get_series("sofr", start, end)

    def get_libor_history(self, tenor_months: int, start: date, end: date) -> pd.DataFrame:
        key = _TENOR_TO_LIBOR_KEY[tenor_months]
        return self.get_series(key, start, end)

    def get_swap_history(self, tenor_years: int, start: date, end: date) -> pd.DataFrame:
        key = _TENOR_TO_SWAP_KEY[tenor_years]
        return self.get_series(key, start, end)

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        """Backup SOFR fixing for ``valuation_date`` as a MarketQuote (FIXING/RATE, decimal)."""
        from datetime import timedelta

        start = valuation_date - timedelta(days=10)
        df = self.get_sofr(start, valuation_date)
        if df.empty:
            return []
        latest = df.iloc[-1]
        return [
            MarketQuote(
                valuation_date=valuation_date,
                instrument_type=InstrumentType.FIXING,
                instrument_id="SOFR",
                maturity_date=latest["date"],
                quote=float(latest["value"]) / 100.0,
                quote_type=QuoteType.RATE,
                source=self.name,
                license=self.license,
                source_url=f"{self.base_url}{_GRAPH_PATH}?id=SOFR",
            )
        ]

    def libor_quotes_from_series(
        self, tenor_months: int, start: date, end: date
    ) -> list[MarketQuote]:
        """Historical LIBOR fixings as MarketQuotes (used for OS-2 Lehman-era reconstruction).
        ``maturity_date`` per row is the fixing date + tenor (approx, 30-day months)."""
        from datetime import timedelta

        df = self.get_libor_history(tenor_months, start, end)
        out = []
        for _, row in df.iterrows():
            fixing_date: date = row["date"]
            maturity = fixing_date + timedelta(days=30 * tenor_months)
            out.append(
                MarketQuote(
                    valuation_date=fixing_date,
                    instrument_type=InstrumentType.FIXING,
                    instrument_id=f"LIBOR_{tenor_months}M",
                    maturity_date=maturity,
                    quote=float(row["value"]) / 100.0,
                    quote_type=QuoteType.RATE,
                    source=self.name,
                    license=self.license,
                    source_url=f"{self.base_url}{_GRAPH_PATH}?id={SERIES_IDS[_TENOR_TO_LIBOR_KEY[tenor_months]]}",
                )
            )
        return out

    def swap_quotes_from_series(
        self, tenor_years: int, start: date, end: date
    ) -> list[MarketQuote]:
        """Historical USD swap par rates as MarketQuotes (used for OS-2 reconstruction)."""
        from datetime import timedelta

        df = self.get_swap_history(tenor_years, start, end)
        out = []
        for _, row in df.iterrows():
            quote_date: date = row["date"]
            maturity = quote_date + timedelta(days=365 * tenor_years)
            out.append(
                MarketQuote(
                    valuation_date=quote_date,
                    instrument_type=InstrumentType.LIBOR_SWAP,
                    instrument_id=f"SWAP_{tenor_years}Y",
                    maturity_date=maturity,
                    quote=float(row["value"]) / 100.0,
                    quote_type=QuoteType.PAR_RATE,
                    source=self.name,
                    license=self.license,
                    source_url=f"{self.base_url}{_GRAPH_PATH}?id={SERIES_IDS[_TENOR_TO_SWAP_KEY[tenor_years]]}",
                )
            )
        return out
