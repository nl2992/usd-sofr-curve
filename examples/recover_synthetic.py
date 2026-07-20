#!/usr/bin/env python
"""Demonstrate dataset OS-1: exact synthetic recovery (docs/PLAN.md §5 step 2, §6).

OS-1 is defined by construction: instruments are built from a KNOWN closed-form reference
discount curve ``P*`` such that every cashflow date (accrual endpoints, coupon dates, futures
period endpoints) coincides with a curve pillar node. That is what makes machine-precision
recovery possible regardless of interpolation scheme — see ``tests/_synthetic.py`` (the
project's canonical OS-1 fixture, reused here rather than duplicated) for exactly how the
instruments are built.

(Off-node cashflows — e.g. a swap coupon date that falls between two bootstrapped pillars —
would instead show a small, expected interpolation-driven discrepancy against ``P*``; that is a
property of the chosen interpolation scheme approximating a smooth curve from finitely many
pillars, not a bootstrap defect. OS-1 sidesteps that entirely by design, which is why it is the
strongest test in the validation ladder: the "true" answer is known and exactly reproducible.)

Runs fully offline. Run with:

    python examples/recover_synthetic.py
"""

from __future__ import annotations

from openusdcurve.curves.bootstrap import bootstrap
from tests._synthetic import (
    NODE_DATES,
    VALUATION_DATE,
    build_synthetic_instruments,
    reference_df,
)


def main() -> None:
    instruments = build_synthetic_instruments()
    curve = bootstrap(instruments, VALUATION_DATE)

    print(f"OS-1 synthetic recovery check — valuation date {VALUATION_DATE}\n")
    print(f"{'node_date':12s} {'P_boot':>18s} {'P_reference':>18s} {'abs_error':>12s}")

    max_err = 0.0
    for d in NODE_DATES:
        p_boot = curve.discount(d)
        p_ref = reference_df(d)
        err = abs(p_boot - p_ref)
        max_err = max(max_err, err)
        print(f"{d.isoformat():12s} {p_boot:18.14f} {p_ref:18.14f} {err:12.3e}")

    print(f"\nmax |P_boot - P_reference| across all nodes = {max_err:.3e}")
    print("PASS (< 1e-9)" if max_err < 1e-9 else "FAIL (>= 1e-9)")


if __name__ == "__main__":
    main()
