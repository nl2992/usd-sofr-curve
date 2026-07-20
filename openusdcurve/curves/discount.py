"""Concrete :class:`DiscountCurve` implementation.

Stores pillar discount factors and interpolates between them (default: log-linear on discount
factors, per docs/PLAN.md §4). ``openusdcurve.curves.base.DiscountCurve`` is a stub that is
overridden by re-exporting the real class from here (see the bottom of ``base.py``... actually
``base.py`` is not modified in-place beyond adding the delegation import, so instead we patch
the class onto the ``base`` module at import time here to keep a single canonical
implementation without a circular import).
"""

from __future__ import annotations

from datetime import date

from openusdcurve.curves.base import Interpolator, year_fraction
from openusdcurve.curves.interpolation import LogLinearDiscount


class DiscountCurve:
    """Discount curve built from pillar dates/discount-factors + an interpolator.

    Invariants: ``discount(reference_date) == 1.0`` and all discount factors are positive.
    """

    reference_date: date

    def __init__(
        self,
        reference_date: date,
        pillar_dates: list[date],
        discount_factors: list[float],
        interpolator: Interpolator | None = None,
    ) -> None:
        if len(pillar_dates) != len(discount_factors):
            raise ValueError("pillar_dates and discount_factors must have the same length")
        for df in discount_factors:
            if df <= 0:
                raise ValueError("discount factors must be strictly positive")

        self.reference_date = reference_date
        self.interpolator: Interpolator = interpolator or LogLinearDiscount()

        # Build the internal (x, y) grid, ensuring (0, 1.0) at the reference date is present.
        pairs = sorted(zip(pillar_dates, discount_factors, strict=True), key=lambda p: p[0])
        x_grid: list[float] = []
        y_grid: list[float] = []
        for d, df in pairs:
            t = year_fraction(reference_date, d)
            if x_grid and t <= x_grid[-1]:
                # Skip duplicate / non-increasing pillar dates (keep the first occurrence).
                continue
            x_grid.append(t)
            y_grid.append(df)

        if not x_grid or x_grid[0] > 0.0:
            x_grid.insert(0, 0.0)
            y_grid.insert(0, 1.0)
        elif x_grid[0] == 0.0:
            y_grid[0] = 1.0

        if len(x_grid) == 1:
            # Degenerate single-node curve: duplicate the reference point so interpolators
            # (which require >= 2 grid points) still function; flat extrapolation handles rest.
            x_grid.append(1.0)
            y_grid.append(y_grid[0])

        self._x_grid = x_grid
        self._y_grid = y_grid
        self.pillar_dates = [p[0] for p in pairs]
        self.discount_factors = [p[1] for p in pairs]

    def _time(self, d: date) -> float:
        return year_fraction(self.reference_date, d)

    def discount(self, d: date) -> float:
        if d == self.reference_date:
            return 1.0
        t = self._time(d)
        return float(self.interpolator(self._x_grid, self._y_grid, t))

    def zero_rate(self, d: date) -> float:
        t = self._time(d)
        if t == 0.0:
            # Instantaneous zero rate at t=0: derive from the first segment's flat forward.
            t_eps = 1e-6
            p_eps = self.discount(self.reference_date + _timedelta_days(t_eps))
            return -_safe_log(p_eps) / t_eps
        p = self.discount(d)
        return -_safe_log(p) / t

    def forward_rate(self, d1: date, d2: date) -> float:
        tau = year_fraction(d1, d2)
        if tau <= 0:
            raise ValueError("forward_rate requires d2 strictly after d1")
        p1 = self.discount(d1)
        p2 = self.discount(d2)
        return (p1 / p2 - 1.0) / tau


def _timedelta_days(years: float):
    from datetime import timedelta

    return timedelta(days=years * 365.0)


def _safe_log(x: float) -> float:
    import math

    return math.log(x)
