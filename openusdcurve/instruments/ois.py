"""SOFR OIS calibration instrument: annual fixed vs daily-compounded SOFR (act/360).

Under single-curve valuation, the daily-compounded-SOFR floating leg telescopes exactly to
``P(effective) - P(maturity)`` (the standard OIS bootstrap identity — compounding a curve's own
overnight forward rates through a period reproduces the discount-factor ratio for that period).
The fixed leg accrues annually on act/360.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import generate_schedule, year_fraction


@dataclass
class OIS:
    """A single-curve SOFR OIS swap: annual fixed (act/360) vs daily-compounded SOFR."""

    instrument_id: str
    valuation_date: date
    effective_date: date
    maturity_date: date
    target_quote: float
    fixed_frequency: str = "1Y"
    fixed_day_count: str = "act360"
    fixed_schedule: list[date] = field(default_factory=list, repr=False)

    def __post_init__(self) -> None:
        if not self.fixed_schedule:
            self.fixed_schedule = generate_schedule(
                self.effective_date, self.maturity_date, self.fixed_frequency
            )

    @property
    def pillar_date(self) -> date:
        return self.maturity_date

    def _annuity(self, curve: Curve) -> float:
        annuity = 0.0
        prev = self.effective_date
        for d in self.fixed_schedule:
            tau = year_fraction(prev, d, self.fixed_day_count)
            annuity += tau * curve.discount(d)
            prev = d
        return annuity

    def implied_quote(self, curve: Curve) -> float:
        p_start = curve.discount(self.effective_date)
        p_end = curve.discount(self.maturity_date)
        annuity = self._annuity(curve)
        return (p_start - p_end) / annuity
