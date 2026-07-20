"""Unit tests for interpolators: identity at nodes, monotonicity, and log-linear correctness."""

from __future__ import annotations

import math

import pytest

from openusdcurve.curves.interpolation import (
    FlatForward,
    Linear,
    LogLinearDiscount,
)

X = [0.0, 1.0, 2.0, 5.0]
DF = [1.0, 0.96, 0.91, 0.80]  # a decreasing discount-factor grid


@pytest.mark.parametrize("interp", [Linear(), FlatForward(), LogLinearDiscount()])
def test_identity_at_nodes(interp):
    for x, y in zip(X, DF, strict=True):
        assert interp(X, DF, x) == pytest.approx(y, abs=1e-12)


def test_linear_midpoint():
    interp = Linear()
    # midpoint between x=0 (1.0) and x=1 (0.96) -> 0.98
    assert interp(X, DF, 0.5) == pytest.approx(0.98)


def test_loglinear_geometric_midpoint():
    interp = LogLinearDiscount()
    # log-linear at midpoint -> geometric mean of the two node DFs.
    expected = math.sqrt(DF[0] * DF[1])
    assert interp(X, DF, 0.5) == pytest.approx(expected)


def test_loglinear_constant_forward():
    # Between two nodes, log-linear implies a constant continuously-compounded forward rate.
    interp = LogLinearDiscount()
    x0, x1 = X[1], X[2]
    y0, y1 = DF[1], DF[2]
    f = (math.log(y0) - math.log(y1)) / (x1 - x0)
    # Check midpoint DF equals y0 * exp(-f * dt)
    xm = 0.5 * (x0 + x1)
    assert interp(X, DF, xm) == pytest.approx(y0 * math.exp(-f * (xm - x0)))


@pytest.mark.parametrize("interp", [Linear(), LogLinearDiscount()])
def test_monotonic_decreasing(interp):
    # A decreasing DF grid should interpolate to a monotonically decreasing curve.
    xs = [i * 0.1 for i in range(0, 51)]
    ys = [interp(X, DF, x) for x in xs]
    assert all(ys[i] >= ys[i + 1] - 1e-12 for i in range(len(ys) - 1))


def test_flatforward_steps():
    interp = FlatForward()
    # Between node 1 (x=1, y=0.96) and node 2 (x=2), value holds at the left node.
    assert interp(X, DF, 1.5) == pytest.approx(DF[1])
    assert interp(X, DF, 1.999) == pytest.approx(DF[1])


def test_loglinear_extrapolation_flat_forward():
    interp = LogLinearDiscount()
    # Beyond last node, forward rate stays constant -> DF keeps decreasing geometrically.
    y_last = interp(X, DF, X[-1])
    y_beyond = interp(X, DF, X[-1] + 1.0)
    assert y_beyond < y_last


def test_positivity_of_loglinear():
    interp = LogLinearDiscount()
    for x in [0.3, 1.7, 4.2, 6.0]:
        assert interp(X, DF, x) > 0.0
