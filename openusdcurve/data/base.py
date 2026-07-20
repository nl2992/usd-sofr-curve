"""Normalized market-data contract.

Every data provider — NY Fed, FRED, CME, Treasury, manual CSV, synthetic — returns the SAME
object: a list of :class:`MarketQuote`. The curve engine never learns which source produced a
quote. Do not change the shape of ``MarketQuote`` or the ``DataSource`` protocol without
updating ``docs/PLAN.md`` §4, as all downstream code depends on them.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from enum import Enum
from typing import Protocol, runtime_checkable


class InstrumentType(str, Enum):
    DEPOSIT = "deposit"
    FUTURE = "future"
    LIBOR_SWAP = "libor_swap"
    OIS = "ois"
    FIXING = "fixing"
    SOFR_INDEX = "sofr_index"
    TREASURY_PAR = "treasury_par"


class QuoteType(str, Enum):
    RATE = "rate"          # simple/annualized rate, decimal (0.0525 == 5.25%)
    PRICE = "price"        # futures price (95.25 == 100 - 4.75)
    PAR_RATE = "par_rate"  # par swap/OIS rate, decimal
    INDEX = "index"        # index level (e.g. SOFR Index)


class LicenseTag(str, Enum):
    PUBLIC_DOMAIN = "public-domain"
    PUBLIC_DISPLAY_ONLY = "public-display-only"
    API_REDISTRIBUTABLE = "api-redistributable"
    MANUAL_USER_SUPPLIED = "manual-user-supplied"
    LICENSED = "licensed"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class MarketQuote:
    """A single normalized observation, source-agnostic.

    Rates are always decimals (``0.0525``), never percents. Futures use ``QuoteType.PRICE``.
    """

    valuation_date: date
    instrument_type: InstrumentType
    instrument_id: str
    maturity_date: date
    quote: float
    quote_type: QuoteType
    source: str
    license: LicenseTag = LicenseTag.UNKNOWN
    observed_at: datetime | None = None
    start_date: date | None = None
    bid: float | None = None
    ask: float | None = None
    source_url: str | None = None
    delayed: bool = False
    provisional: bool = False


@runtime_checkable
class DataSource(Protocol):
    """Every provider implements this. Return a normalized, ready-to-bootstrap quote list."""

    name: str
    license: LicenseTag

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        ...
