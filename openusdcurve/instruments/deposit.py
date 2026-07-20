"""Cash deposit calibration instrument (LIBOR/SOFR-fixing style, act/360 by default)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import year_fraction


@dataclass
class Deposit:
    """A single cash deposit: simple rate accruing from ``start_date`` to ``maturity_date``.

    ``target_quote`` is the deposit's simple rate (decimal, e.g. 0.0525).
    """

    instrument_id: str
    valuation_date: date
    start_date: date
    maturity_date: date
    target_quote: float
    day_count: str = "act360"

    @property
    def pillar_date(self) -> date:
        return self.maturity_date

    def implied_quote(self, curve: Curve) -> float:
        tau = year_fraction(self.start_date, self.maturity_date, self.day_count)
        p_start = curve.discount(self.start_date)
        p_end = curve.discount(self.maturity_date)
        return (p_start / p_end - 1.0) / tau
