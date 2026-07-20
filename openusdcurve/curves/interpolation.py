"""Interpolators matching the ``Interpolator`` protocol in ``curves/base.py``.

Each implementation is a callable ``(x_grid, y_grid, x) -> float`` operating on a generic
ordered grid. ``DiscountCurve`` uses these with x = time (year fraction) and y = discount
factor. The project default is :class:`LogLinearDiscount` (linear in log-DF, i.e. piecewise-
constant instantaneous forward rates), per docs/PLAN.md §4.

All interpolators flat-extrapolate beyond the grid edges: the value at the boundary is held
constant outside [x_grid[0], x_grid[-1]] for ``Linear`` and ``FlatForward``. ``LogLinearDiscount``
extrapolates by carrying forward the last segment's continuously-compounded rate (i.e. flat
forward rate beyond the last pillar, and flat back to the first).
"""

from __future__ import annotations

import bisect
import math


def _locate(x_grid: list[float], x: float) -> int:
    """Return index ``i`` such that x_grid[i] <= x <= x_grid[i+1], clamped to valid range.

    Assumes x_grid is sorted ascending with at least 2 points.
    """
    n = len(x_grid)
    if n < 2:
        raise ValueError("x_grid must have at least 2 points")
    i = bisect.bisect_right(x_grid, x) - 1
    if i < 0:
        i = 0
    if i > n - 2:
        i = n - 2
    return i


class Linear:
    """Linear interpolation on (x, y), flat-extrapolated beyond the grid."""

    def __call__(self, x_grid: list[float], y_grid: list[float], x: float) -> float:
        if len(x_grid) == 1:
            return y_grid[0]
        if x <= x_grid[0]:
            if x == x_grid[0]:
                return y_grid[0]
            return y_grid[0]
        if x >= x_grid[-1]:
            return y_grid[-1]
        i = _locate(x_grid, x)
        x0, x1 = x_grid[i], x_grid[i + 1]
        y0, y1 = y_grid[i], y_grid[i + 1]
        w = (x - x0) / (x1 - x0)
        return y0 + w * (y1 - y0)


class FlatForward:
    """Piecewise-constant (step) interpolation: holds the left node's value across each segment."""

    def __call__(self, x_grid: list[float], y_grid: list[float], x: float) -> float:
        if len(x_grid) == 1:
            return y_grid[0]
        if x <= x_grid[0]:
            return y_grid[0]
        if x >= x_grid[-1]:
            return y_grid[-1]
        i = _locate(x_grid, x)
        return y_grid[i]


class LogLinearDiscount:
    """Linear interpolation in log(y) vs x — the standard log-linear-on-discount-factor scheme.

    Equivalent to piecewise-constant continuously-compounded forward rates between pillars.
    y_grid values must be strictly positive (discount factors).
    """

    def __call__(self, x_grid: list[float], y_grid: list[float], x: float) -> float:
        if len(x_grid) == 1:
            return y_grid[0]

        n = len(x_grid)
        if x < x_grid[0]:
            # Flat-extrapolate the first segment's rate back to x=0 / before the first pillar.
            i = 0
        elif x > x_grid[-1]:
            # Flat-extrapolate the last segment's rate beyond the final pillar.
            i = n - 2
        else:
            i = _locate(x_grid, x)

        x0, x1 = x_grid[i], x_grid[i + 1]
        y0, y1 = y_grid[i], y_grid[i + 1]
        log_y0, log_y1 = math.log(y0), math.log(y1)
        if x1 == x0:
            return y0
        w = (x - x0) / (x1 - x0)
        log_y = log_y0 + w * (log_y1 - log_y0)
        return math.exp(log_y)


DEFAULT_INTERPOLATOR = LogLinearDiscount
