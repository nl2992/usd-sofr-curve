"""NY Fed Markets Data API — SOFR fixings and SOFR Index (authoritative public source).

API assumptions (documented, not independently verified against a live call in this repo — see
TODO below). The NY Fed publishes a "Reference Rates" search endpoint returning daily secured
overnight rates (SOFR/BGCR/TGCR) and a related "SOFR Averages and Index" endpoint. This module
targets the shapes:

- Rates search:  GET {base_url}/rates/secured/sofr/search.json?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD
  -> {"refRates": [{"effectiveDate": "2026-07-20", "type": "SOFR", "percentRate": 5.31,
                     "percentPercentile1": 5.28, "percentPercentile25": 5.30,
                     "percentPercentile75": 5.33, "percentPercentile99": 5.40,
                     "volumeInBillions": 1950.0, "revisionIndicator": "N"}, ...]}

- SOFR Index/averages search: GET {base_url}/rates/secured/sofrai/search.json?startDate=...&endDate=...
  -> {"refRates": [{"effectiveDate": "2026-07-20", "index": 1.14872341,
                     "sofr30DayAverage": 5.29, "sofr90DayAverage": 5.31,
                     "sofr180DayAverage": 5.20, "revisionIndicator": "N"}, ...]}

TODO: verify exact field names/paths against a live response before relying on this in
production; the NY Fed API has changed field names historically (e.g. `percentRate` vs `rate`).
This module is written so only ``_RATES_PATH``/``_INDEX_PATH`` and the small parsing helpers
need to change if the live schema differs — the public method signatures should not.

Design constraints (per task spec): importing this module must never perform network I/O, and
the HTTP client must be injectable so tests can run fully offline against committed sample JSON
(``data/sample/nyfed_sofr_sample.json``, ``data/sample/nyfed_sofr_index_sample.json``).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["NewYorkFedSOFRSource"]

_BASE_URL = "https://markets.newyorkfed.org/api"
_RATES_PATH = "/rates/secured/sofr/search.json"
_INDEX_PATH = "/rates/secured/sofrai/search.json"


def _parse_rates_json(payload: dict) -> pd.DataFrame:
    rows = payload.get("refRates", [])
    if not rows:
        return pd.DataFrame(
            columns=["effective_date", "rate", "volume_billions", "p1", "p25", "p75", "p99"]
        )
    out = []
    for r in rows:
        out.append(
            {
                "effective_date": date.fromisoformat(r["effectiveDate"]),
                "rate": float(r["percentRate"]) / 100.0,  # normalize percent -> decimal
                "volume_billions": r.get("volumeInBillions"),
                "p1": _pct(r.get("percentPercentile1")),
                "p25": _pct(r.get("percentPercentile25")),
                "p75": _pct(r.get("percentPercentile75")),
                "p99": _pct(r.get("percentPercentile99")),
            }
        )
    df = pd.DataFrame(out).sort_values("effective_date").reset_index(drop=True)
    return df


def _pct(v: float | None) -> float | None:
    return None if v is None else float(v) / 100.0


def _parse_index_json(payload: dict) -> pd.DataFrame:
    rows = payload.get("refRates", [])
    if not rows:
        return pd.DataFrame(
            columns=["effective_date", "index", "avg_30d", "avg_90d", "avg_180d"]
        )
    out = []
    for r in rows:
        out.append(
            {
                "effective_date": date.fromisoformat(r["effectiveDate"]),
                "index": float(r["index"]),
                "avg_30d": _pct(r.get("sofr30DayAverage")),
                "avg_90d": _pct(r.get("sofr90DayAverage")),
                "avg_180d": _pct(r.get("sofr180DayAverage")),
            }
        )
    df = pd.DataFrame(out).sort_values("effective_date").reset_index(drop=True)
    return df


@dataclass
class NewYorkFedSOFRSource:
    """SOFR fixing + SOFR Index source. Network calls are lazy: nothing happens at import or
    construction time, only when ``get_fixings``/``get_sofr_index``/``get_quotes`` is called.

    ``client`` is injectable so tests can pass an ``httpx.Client`` built with
    ``httpx.MockTransport`` instead of hitting the network.
    """

    base_url: str = _BASE_URL
    client: object | None = None  # httpx.Client, created lazily if not supplied
    name: str = "new-york-fed"
    license: LicenseTag = LicenseTag.PUBLIC_DOMAIN

    _owns_client: bool = field(default=False, init=False, repr=False)

    def _get_client(self):
        if self.client is None:
            import httpx  # local import: keep module import side-effect free

            self.client = httpx.Client(base_url=self.base_url, timeout=30.0)
            self._owns_client = True
        return self.client

    def _get_json(self, path: str, params: dict) -> dict:
        client = self._get_client()
        response = client.get(path, params=params)
        response.raise_for_status()
        return response.json()

    def get_fixings(self, start: date, end: date) -> pd.DataFrame:
        """Daily SOFR fixings (and percentile band + volume) for [start, end]."""
        payload = self._get_json(
            _RATES_PATH, {"startDate": start.isoformat(), "endDate": end.isoformat()}
        )
        return _parse_rates_json(payload)

    def get_sofr_index(self, start: date, end: date) -> pd.DataFrame:
        """SOFR Index level + 30/90/180-day averages for [start, end]."""
        payload = self._get_json(
            _INDEX_PATH, {"startDate": start.isoformat(), "endDate": end.isoformat()}
        )
        return _parse_index_json(payload)

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        """SOFR fixing + SOFR Index as MarketQuotes for ``valuation_date`` (looks back a few
        days in case the target date has no published fixing yet, e.g. same-day query)."""
        from datetime import timedelta

        lookback_start = valuation_date - timedelta(days=7)
        quotes: list[MarketQuote] = []

        fixings = self.get_fixings(lookback_start, valuation_date)
        if not fixings.empty:
            latest = fixings.iloc[-1]
            quotes.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.FIXING,
                    instrument_id="SOFR",
                    maturity_date=latest["effective_date"],
                    quote=float(latest["rate"]),
                    quote_type=QuoteType.RATE,
                    source=self.name,
                    license=self.license,
                    observed_at=datetime.combine(latest["effective_date"], datetime.min.time()),
                    source_url=f"{self.base_url}{_RATES_PATH}",
                )
            )

        index_df = self.get_sofr_index(lookback_start, valuation_date)
        if not index_df.empty:
            latest_idx = index_df.iloc[-1]
            quotes.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.SOFR_INDEX,
                    instrument_id="SOFR_INDEX",
                    maturity_date=latest_idx["effective_date"],
                    quote=float(latest_idx["index"]),
                    quote_type=QuoteType.INDEX,
                    source=self.name,
                    license=self.license,
                    observed_at=datetime.combine(
                        latest_idx["effective_date"], datetime.min.time()
                    ),
                    source_url=f"{self.base_url}{_INDEX_PATH}",
                )
            )
        return quotes


def load_sample_rates_json(path: str) -> pd.DataFrame:
    """Test/offline helper: parse a committed sample rates JSON file without network I/O."""
    with open(path, encoding="utf-8") as fh:
        payload = json.load(fh)
    return _parse_rates_json(payload)


def load_sample_index_json(path: str) -> pd.DataFrame:
    """Test/offline helper: parse a committed sample SOFR-index JSON file without network I/O."""
    with open(path, encoding="utf-8") as fh:
        payload = json.load(fh)
    return _parse_index_json(payload)
