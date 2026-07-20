"""Daily-compounded SOFR (OIS-style) growth factors, from fixings or from a curve.

Implements the NY Fed SOFR Index relationship:

    Index(t2) / Index(t1) = prod_i (1 + SOFR_i * n_i / 360)

where the product runs over each US business day ``i`` in ``[t1, t2)`` with a published daily
fixing, and ``n_i`` is the number of calendar days that fixing applies to (1 on an ordinary day;
3 for a Friday fixing that also covers Saturday/Sunday; more across a long weekend/holiday).
This is exactly the act/360 daily-compounding convention used by the ``OIS`` calibration
instrument (``openusdcurve/instruments/ois.py``), so pricing here is consistent with bootstrap.

Two ways to build the growth factor:

- ``compound_factor_from_fixings`` — from a realized daily fixings series (validation ladder L3
  recomputes the published NY Fed SOFR Index this way).
- ``projected_compound_factor`` — from a curve's own overnight forward rates (used implicitly by
  the ``OIS`` instrument's single-curve telescoping identity: compounding a curve's own daily
  forwards over a period reproduces ``P(start)/P(end)`` exactly, regardless of compounding
  granularity).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import DEFAULT_CALENDAR, USBusinessCalendar

__all__ = [
    "Fixing",
    "business_days_in_range",
    "compound_factor_from_fixings",
    "compound",
    "projected_compound_factor",
    "projected_compound_rate",
]


@dataclass(frozen=True)
class Fixing:
    """A single published daily SOFR fixing."""

    effective_date: date
    rate: float


def business_days_in_range(
    start: date, end: date, calendar: USBusinessCalendar | None = None
) -> list[date]:
    """US business days in ``[start, end)`` (start rolled forward if it isn't one)."""
    cal = calendar or DEFAULT_CALENDAR
    d = start if cal.is_business_day(start) else cal.next_business_day(start)
    days: list[date] = []
    while d < end:
        days.append(d)
        d = cal.next_business_day(d)
    return days


def compound_factor_from_fixings(
    fixings: list[Fixing],
    start: date,
    end: date,
    calendar: USBusinessCalendar | None = None,
) -> float:
    """Growth factor ``prod_i (1 + r_i * n_i/360)`` over ``[start, end)`` from daily fixings.

    ``n_i`` is the calendar-day span from business day ``i`` to the next business day (or to
    ``end`` for the final one), matching the NY Fed SOFR Index compounding convention.
    """
    cal = calendar or DEFAULT_CALENDAR
    rates = {f.effective_date: f.rate for f in fixings}
    days = business_days_in_range(start, end, cal)
    if not days:
        return 1.0

    factor = 1.0
    for i, d in enumerate(days):
        if d not in rates:
            raise ValueError(f"Missing SOFR fixing for business day {d}")
        nxt = days[i + 1] if i + 1 < len(days) else end
        n = (nxt - d).days
        factor *= 1.0 + rates[d] * n / 360.0
    return factor


def compound(
    fixings: list[Fixing],
    start: date,
    end: date,
    calendar: USBusinessCalendar | None = None,
) -> float:
    """Daily-compounded simple SOFR rate (act/360) over ``[start, end)`` from fixings."""
    factor = compound_factor_from_fixings(fixings, start, end, calendar)
    tau = (end - start).days / 360.0
    if tau <= 0:
        raise ValueError("compound requires end strictly after start")
    return (factor - 1.0) / tau


def projected_compound_factor(
    curve: Curve,
    start: date,
    end: date,
    calendar: USBusinessCalendar | None = None,
) -> float:
    """Growth factor over ``[start, end)`` from a curve's own act/360 overnight forwards.

    Because each daily factor is ``1 + fwd_i * tau_i`` with ``fwd_i = (P(d)/P(next_d) - 1) /
    tau_i``, the product telescopes exactly to ``curve.discount(start) / curve.discount(end)`` —
    the same identity the ``OIS`` instrument relies on for single-curve valuation. This function
    makes that identity explicit and reusable for reporting/validation.
    """
    cal = calendar or DEFAULT_CALENDAR
    days = business_days_in_range(start, end, cal)
    if not days:
        return curve.discount(start) / curve.discount(end)

    factor = 1.0
    for i, d in enumerate(days):
        nxt = days[i + 1] if i + 1 < len(days) else end
        tau = (nxt - d).days / 360.0
        p1 = curve.discount(d)
        p2 = curve.discount(nxt)
        fwd = (p1 / p2 - 1.0) / tau
        factor *= 1.0 + fwd * tau
    return factor


def projected_compound_rate(
    curve: Curve,
    start: date,
    end: date,
    calendar: USBusinessCalendar | None = None,
) -> float:
    """Daily-compounded simple rate (act/360) over ``[start, end)`` projected from ``curve``."""
    factor = projected_compound_factor(curve, start, end, calendar)
    tau = (end - start).days / 360.0
    if tau <= 0:
        raise ValueError("projected_compound_rate requires end strictly after start")
    return (factor - 1.0) / tau
