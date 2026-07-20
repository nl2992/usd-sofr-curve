"""Compounding a constant SOFR reproduces analytic (1+r*n/360)-style index growth, matches a
known fixtures series, and (via the curve) reproduces the OIS single-curve telescoping identity.
"""

from __future__ import annotations

from datetime import date

import pytest

from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.pricing.sofr_compounding import (
    Fixing,
    business_days_in_range,
    compound,
    compound_factor_from_fixings,
    projected_compound_factor,
    projected_compound_rate,
)
from tests._synthetic import VALUATION_DATE, build_synthetic_instruments


def test_constant_rate_matches_analytic_growth_across_a_weekend():
    # 2025-01-06 is a Monday; +7 calendar days lands on the following Monday, crossing one
    # weekend so the Friday fixing applies for n=3 calendar days.
    start = date(2025, 1, 6)
    end = date(2025, 1, 13)
    r = 0.0525

    days = business_days_in_range(start, end)
    assert days == [date(2025, 1, 6), date(2025, 1, 7), date(2025, 1, 8), date(2025, 1, 9), date(2025, 1, 10)]

    fixings = [Fixing(effective_date=d, rate=r) for d in days]
    factor = compound_factor_from_fixings(fixings, start, end)

    expected = (1.0 + r * 1 / 360.0) ** 4 * (1.0 + r * 3 / 360.0)
    assert factor == pytest.approx(expected, rel=1e-12)

    # compound() returns the equivalent simple act/360 rate over the whole period.
    equiv_rate = compound(fixings, start, end)
    tau = (end - start).days / 360.0
    assert factor == pytest.approx(1.0 + equiv_rate * tau, rel=1e-12)


def test_matches_known_fixtures_series_no_weekend():
    # Mon/Tue/Wed with Thu as the (exclusive) end date: n=1 for every fixing.
    start = date(2025, 1, 6)
    end = date(2025, 1, 9)
    rates = [0.0510, 0.0512, 0.0508]
    fixings = [
        Fixing(effective_date=date(2025, 1, 6), rate=rates[0]),
        Fixing(effective_date=date(2025, 1, 7), rate=rates[1]),
        Fixing(effective_date=date(2025, 1, 8), rate=rates[2]),
    ]

    factor = compound_factor_from_fixings(fixings, start, end)
    expected = 1.0
    for r in rates:
        expected *= 1.0 + r * 1 / 360.0

    assert factor == pytest.approx(expected, rel=1e-12)


def test_missing_fixing_raises_value_error():
    start = date(2025, 1, 6)
    end = date(2025, 1, 9)
    fixings = [Fixing(effective_date=date(2025, 1, 6), rate=0.05)]  # Tue, Wed missing
    with pytest.raises(ValueError):
        compound_factor_from_fixings(fixings, start, end)


def test_empty_range_returns_unit_factor():
    d = date(2025, 1, 6)
    assert compound_factor_from_fixings([], d, d) == 1.0


def test_projected_compound_factor_matches_curve_discount_ratio():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    from openusdcurve.instruments.conventions import add_tenor

    start = add_tenor(VALUATION_DATE, "6M")
    end = add_tenor(VALUATION_DATE, "12M")

    factor = projected_compound_factor(boot, start, end)
    expected = boot.discount(start) / boot.discount(end)
    assert factor == pytest.approx(expected, rel=1e-10)

    rate = projected_compound_rate(boot, start, end)
    tau = (end - start).days / 360.0
    assert factor == pytest.approx(1.0 + rate * tau, rel=1e-12)
