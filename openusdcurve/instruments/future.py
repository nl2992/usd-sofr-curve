"""Eurodollar / SOFR future calibration instrument.

Price = 100 * (1 - forward_rate) over the future's reference period (typically 3M IMM).
No convexity adjustment is applied here (TODO below): a full implementation would subtract a
convexity correction from the model forward rate before converting to price, growing with
tenor^2 and rate volatility.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import year_fraction


@dataclass
class Future:
    """A single futures contract implying a forward rate over [period_start, period_end].

    ``target_quote`` is the futures settlement price (e.g. 95.25 for a 4.75% implied rate).
    """

    instrument_id: str
    valuation_date: date
    period_start: date
    period_end: date
    target_quote: float
    day_count: str = "act360"
    convexity_adjustment: float = 0.0  # TODO: replace with a model-based convexity correction

    @property
    def pillar_date(self) -> date:
        return self.period_end

    def implied_quote(self, curve: Curve) -> float:
        tau = year_fraction(self.period_start, self.period_end, self.day_count)
        p_start = curve.discount(self.period_start)
        p_end = curve.discount(self.period_end)
        fwd_rate = (p_start / p_end - 1.0) / tau
        # TODO: fwd_rate -= self._convexity_adjustment(curve) for futures-vs-FRA convexity bias.
        fwd_rate -= self.convexity_adjustment
        return 100.0 * (1.0 - fwd_rate)
