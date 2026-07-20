"""CME public delayed SOFR futures settlements — a REPLACEABLE ADAPTER.

CME's public (non-subscriber) delayed quote pages/settlement files are >=10 minutes delayed and
carry redistribution restrictions (PLAN §3: license = public-display-only). This module
deliberately does NOT embed a scraper as a hard dependency: scraping CME's site is fragile
(subject to markup/ToS changes) and out of scope for release 0.1 (PLAN §10: "Web scraping is
NOT a prerequisite").

Two supported paths:

1. ``from_csv`` — parse a manually downloaded CME settlement file (the common path for now).
2. ``scraper`` hook — inject any callable ``(valuation_date) -> list[MarketQuote]`` at
   construction time to plug in a live scraper or vendor client later, without touching this
   adapter's interface. See the ``# HOOK`` marker below.

Expected settlement CSV schema (columns, case-insensitive; matches typical CME "End-of-Day
Settlement" exports for SOFR futures — SR1/SR3):

- ``contract_code``   e.g. ``SR3U26`` (product + month code + year).
- ``contract_month``  ISO date of the contract month/expiry reference, e.g. ``2026-09-16``.
- ``settle_price``    Futures settlement price (e.g. 94.62 == 100 - 5.38%).
- ``volume``          Optional, contract volume traded.
- ``open_interest``   Optional, open interest.
- ``trade_date``      ISO date the settlement applies to.

An example file ships at ``data/sample/cme_settlement_sample.csv``.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["CMEPublicDelayedSource"]

_SETTLEMENT_REQUIRED_COLUMNS = (
    "contract_code",
    "contract_month",
    "settle_price",
    "trade_date",
)


def _parse_settlement_csv(path: str | Path, valuation_date: date | None = None) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = [c.strip().lower() for c in df.columns]
    missing = [c for c in _SETTLEMENT_REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{path}: missing required column(s) {missing}; "
            f"required={_SETTLEMENT_REQUIRED_COLUMNS}"
        )
    df["contract_month"] = pd.to_datetime(df["contract_month"]).dt.date
    df["trade_date"] = pd.to_datetime(df["trade_date"]).dt.date
    if valuation_date is not None:
        df = df[df["trade_date"] == valuation_date]
    return df.reset_index(drop=True)


@dataclass
class CMEPublicDelayedSource:
    """Adapter for CME SOFR futures settlements. ``delayed=True`` and
    ``license=PUBLIC_DISPLAY_ONLY`` unconditionally, per PLAN §3 (ToS-restricted redistribution).
    """

    csv_path: str | Path | None = None
    scraper: Callable[[date], list[MarketQuote]] | None = None
    name: str = "cme-public-delayed"
    license: LicenseTag = LicenseTag.PUBLIC_DISPLAY_ONLY
    delayed: bool = True

    @classmethod
    def from_csv(cls, path: str | Path, **kwargs) -> CMEPublicDelayedSource:
        return cls(csv_path=path, **kwargs)

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        if self.scraper is not None:
            # HOOK: plug a live scraper / vendor client here. It must return
            # list[MarketQuote] already tagged with this adapter's license/delayed
            # semantics (or override them below) — this adapter does not implement
            # any scraping logic itself.
            quotes = self.scraper(valuation_date)
            return [self._ensure_license(q) for q in quotes]

        if self.csv_path is not None:
            return self._quotes_from_csv(valuation_date)

        raise ValueError(
            "CMEPublicDelayedSource requires either csv_path (manual settlement file) "
            "or a scraper callable; neither was provided."
        )

    def _ensure_license(self, q: MarketQuote) -> MarketQuote:
        # Public-display-only / delayed status is a property of this source, not of whatever
        # produced the raw quote, so it is enforced here regardless of what the scraper set.
        from dataclasses import replace

        return replace(q, license=self.license, delayed=True, source=self.name)

    def _quotes_from_csv(self, valuation_date: date) -> list[MarketQuote]:
        df = _parse_settlement_csv(self.csv_path, valuation_date=valuation_date)
        out = []
        for _, row in df.iterrows():
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.FUTURE,
                    instrument_id=str(row["contract_code"]),
                    maturity_date=row["contract_month"],
                    quote=float(row["settle_price"]),
                    quote_type=QuoteType.PRICE,
                    source=self.name,
                    license=self.license,
                    delayed=True,
                    source_url="https://www.cmegroup.com/markets/interest-rates/stirs/sofr.html",
                )
            )
        return out
