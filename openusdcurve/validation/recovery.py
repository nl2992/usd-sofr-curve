"""Ladder L2 — synthetic curve recovery (docs/PLAN.md §5.2, dataset OS-1).

Wraps the OS-1 idea (build quotes from a known reference curve, bootstrap, compare) into a
reusable, non-raising report function so it can be called outside the test suite (e.g. from a
CLI ``validate --dataset os1`` path) as well as from ``tests/test_bootstrap_recovery.py``-style
tests.
"""

from __future__ import annotations

from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.instruments.base import Instrument
from openusdcurve.validation.report import ValidationItem, ValidationReport

__all__ = ["synthetic_recovery_report"]


def synthetic_recovery_report(
    instruments: list[Instrument],
    valuation_date: date,
    reference_curve: Curve,
    node_dates: list[date] | None = None,
    tolerance: float = 1e-9,
) -> ValidationReport:
    """Bootstrap ``instruments`` and compare recovered discount factors to ``reference_curve``.

    ``node_dates`` defaults to the sorted, de-duplicated set of instrument pillar dates. Never
    raises: a bootstrap failure is reported as a single failing item rather than propagated.
    """
    report = ValidationReport(label="L2 synthetic recovery")

    try:
        boot = bootstrap(instruments, valuation_date)
    except Exception as exc:  # pragma: no cover - defensive; report, don't raise
        report.items.append(
            ValidationItem("bootstrap", "fail", f"bootstrap raised: {exc!r}")
        )
        return report

    dates = node_dates if node_dates is not None else sorted({inst.pillar_date for inst in instruments})

    max_err = 0.0
    for d in dates:
        boot_df = boot.discount(d)
        ref_df = reference_curve.discount(d)
        err = abs(boot_df - ref_df)
        max_err = max(max_err, err)
        status = "pass" if err < tolerance else "fail"
        report.items.append(
            ValidationItem(
                f"recovery.{d.isoformat()}",
                status,
                f"|P_boot - P*| = {err:.3e}",
                {"boot_df": boot_df, "ref_df": ref_df, "error": err},
            )
        )

    report.items.append(
        ValidationItem(
            "recovery.max_error",
            "pass" if max_err < tolerance else "fail",
            f"max |P_boot - P*| at nodes = {max_err:.3e}",
            {"max_error": max_err, "tolerance": tolerance},
        )
    )
    return report
