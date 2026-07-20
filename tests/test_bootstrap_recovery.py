"""OS-1 exact synthetic recovery test (docs/PLAN.md §5.2, the strongest test).

Build a known reference discount curve P*(0,T), generate deposit/future/swap/OIS quotes
consistent with it, bootstrap, and assert the bootstrapped curve reproduces P* at the
calibration nodes to machine precision.
"""

from __future__ import annotations

import pytest

from openusdcurve.curves.bootstrap import bootstrap
from tests._synthetic import (
    NODE_DATES,
    VALUATION_DATE,
    build_synthetic_instruments,
    reference_df,
    true_curve,
)


def test_bootstrap_recovers_reference_curve_at_nodes():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    max_err = 0.0
    for d in NODE_DATES:
        err = abs(boot.discount(d) - reference_df(d))
        max_err = max(max_err, err)

    assert max_err < 1e-9, f"max |P_boot - P*| at nodes = {max_err:.3e}"


def test_bootstrap_reference_date_df_is_one():
    boot = bootstrap(build_synthetic_instruments(), VALUATION_DATE)
    assert boot.discount(VALUATION_DATE) == pytest.approx(1.0, abs=1e-15)


def test_bootstrap_discount_factors_positive_and_decreasing():
    boot = bootstrap(build_synthetic_instruments(), VALUATION_DATE)
    dfs = [boot.discount(d) for d in NODE_DATES]
    assert all(df > 0 for df in dfs)
    assert all(dfs[i] > dfs[i + 1] for i in range(len(dfs) - 1))


def test_bootstrap_reprices_every_input_quote():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    for inst in instruments:
        implied = inst.implied_quote(boot)
        assert implied == pytest.approx(inst.target_quote, abs=1e-10), (
            f"{inst.instrument_id}: implied={implied} target={inst.target_quote}"
        )


def test_true_curve_matches_reference_at_nodes():
    curve = true_curve()
    for d in NODE_DATES:
        assert curve.discount(d) == pytest.approx(reference_df(d), abs=1e-15)
