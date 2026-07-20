"""USD LIBOR par swap calibration instrument: semiannual 30/360 fixed vs 3M act/360 float.

Single-curve valuation (Track A, Lehman-era replication): the same curve discounts and
forwards, so the floating leg telescopes to ``P(effective) - P(maturity)`` and the par rate is
the standard ``(P(t0) - P(tN)) / annuity`` formula, with the fixed-leg annuity computed on the
fixed schedule/day-count.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import generate_schedule, year_fraction


@dataclass
class SwapLIBOR:
    """A single-curve USD LIBOR par swap.

    ``target_quote`` is the par swap rate (decimal). Fixed leg: semiannual, 30/360.
    Float leg: quarterly (3M LIBOR), act/360 — used only implicitly via single-curve telescoping.
    """

    instrument_id: str
    valuation_date: date
    effective_date: date
    maturity_date: date
    target_quote: float
    fixed_frequency: str = "6M"
    fixed_day_count: str = "thirty360"
    float_frequency: str = "3M"
    float_day_count: str = "act360"
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
