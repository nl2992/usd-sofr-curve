"""Ladder L5 — QuantLib comparison (docs/PLAN.md §5.5). OPTIONAL dependency.

``QuantLib`` is not a hard requirement of this project. If ``import QuantLib`` fails, every
function here still exists and returns a :class:`~openusdcurve.validation.report.ValidationReport`
with ``skipped=True`` instead of raising — callers (including the CLI) can always call these
functions unconditionally.
"""

from __future__ import annotations

from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.validation.report import ValidationItem, ValidationReport

__all__ = [
    "HAS_QUANTLIB",
    "compare_discount_factors",
    "compare_zero_rates",
    "compare_par_rates",
]

try:
    import QuantLib as ql  # type: ignore

    HAS_QUANTLIB = True
except ImportError:  # pragma: no cover - exercised whenever QuantLib isn't installed
    ql = None
    HAS_QUANTLIB = False


def _skipped_report(label: str) -> ValidationReport:
    report = ValidationReport(label=label, skipped=True)
    report.items.append(
        ValidationItem("quantlib", "skipped", "QuantLib is not installed; L5 comparison skipped")
    )
    return report


def _to_ql_date(d: date):
    return ql.Date(d.day, d.month, d.year)


def _build_ql_curve(reference_date: date, pillar_dates: list[date], discount_factors: list[float]):
    ql_ref = _to_ql_date(reference_date)
    ql.Settings.instance().evaluationDate = ql_ref
    dates = [ql_ref] + [_to_ql_date(d) for d in pillar_dates]
    dfs = [1.0] + [float(x) for x in discount_factors]
    return ql.DiscountCurve(dates, dfs, ql.Actual365Fixed())


def _curve_nodes(curve: Curve, dates: list[date] | None) -> tuple[list[date], list[float]]:
    pillar_dates = getattr(curve, "pillar_dates", None)
    discount_factors = getattr(curve, "discount_factors", None)
    if pillar_dates is None or discount_factors is None:
        pillar_dates = dates or []
        discount_factors = [curve.discount(d) for d in pillar_dates]
    return pillar_dates, discount_factors


def compare_discount_factors(
    curve: Curve, dates: list[date], tolerance: float = 1e-6
) -> ValidationReport:
    """Compare ``curve``'s discount factors at ``dates`` to a QuantLib rebuild of the same nodes."""
    if not HAS_QUANTLIB:
        return _skipped_report("L5 QuantLib comparison (discount factors)")

    report = ValidationReport(label="L5 QuantLib comparison (discount factors)")
    try:
        pillar_dates, discount_factors = _curve_nodes(curve, dates)
        ql_curve = _build_ql_curve(curve.reference_date, pillar_dates, discount_factors)
    except Exception as exc:  # pragma: no cover - defensive; report, don't raise
        report.items.append(ValidationItem("quantlib.build", "fail", f"failed to build QuantLib curve: {exc!r}"))
        return report

    max_err = 0.0
    for d in dates:
        ours = curve.discount(d)
        theirs = ql_curve.discount(_to_ql_date(d))
        err = abs(ours - theirs)
        max_err = max(max_err, err)
        status = "pass" if err < tolerance else "warn"
        report.items.append(
            ValidationItem(f"df.{d.isoformat()}", status, f"|ours - ql| = {err:.3e}", {"ours": ours, "quantlib": theirs})
        )

    report.items.append(
        ValidationItem("df.max_error", "pass" if max_err < tolerance else "warn", f"max |ours - ql| = {max_err:.3e}")
    )
    return report


def compare_zero_rates(
    curve: Curve, dates: list[date], tolerance: float = 1e-4
) -> ValidationReport:
    """Compare ``curve``'s (act/365f continuously-compounded) zero rates to QuantLib's."""
    if not HAS_QUANTLIB:
        return _skipped_report("L5 QuantLib comparison (zero rates)")

    report = ValidationReport(label="L5 QuantLib comparison (zero rates)")
    try:
        pillar_dates, discount_factors = _curve_nodes(curve, dates)
        ql_curve = _build_ql_curve(curve.reference_date, pillar_dates, discount_factors)
    except Exception as exc:  # pragma: no cover
        report.items.append(ValidationItem("quantlib.build", "fail", f"failed to build QuantLib curve: {exc!r}"))
        return report

    day_counter = ql.Actual365Fixed()
    max_err = 0.0
    for d in dates:
        ours = curve.zero_rate(d)
        theirs = ql_curve.zeroRate(_to_ql_date(d), day_counter, ql.Continuous).rate()
        err = abs(ours - theirs)
        max_err = max(max_err, err)
        status = "pass" if err < tolerance else "warn"
        report.items.append(
            ValidationItem(f"zero.{d.isoformat()}", status, f"|ours - ql| = {err:.3e}", {"ours": ours, "quantlib": theirs})
        )

    report.items.append(
        ValidationItem("zero.max_error", "pass" if max_err < tolerance else "warn", f"max |ours - ql| = {max_err:.3e}")
    )
    return report


def compare_par_rates(
    curve: Curve,
    dates: list[date],
    tenor_years: float = 1.0,
    tolerance: float = 1e-4,
) -> ValidationReport:
    """Compare simple forward rates over ``tenor_years`` (a lightweight par-rate proxy) to QuantLib.

    This is not a full swap-schedule par-rate comparison (that needs QuantLib's swap-index/leg
    machinery and is out of scope for a dependency-optional check); it compares the same simple
    forward-rate identity both engines agree on: ``(P(d1)/P(d2) - 1) / tau``.
    """
    if not HAS_QUANTLIB:
        return _skipped_report("L5 QuantLib comparison (forward/par proxy rates)")

    report = ValidationReport(label="L5 QuantLib comparison (forward/par proxy rates)")
    try:
        pillar_dates, discount_factors = _curve_nodes(curve, dates)
        ql_curve = _build_ql_curve(curve.reference_date, pillar_dates, discount_factors)
    except Exception as exc:  # pragma: no cover
        report.items.append(ValidationItem("quantlib.build", "fail", f"failed to build QuantLib curve: {exc!r}"))
        return report

    from openusdcurve.instruments.conventions import add_tenor

    day_counter = ql.Actual360()
    max_err = 0.0
    for d in dates:
        d2 = add_tenor(d, f"{int(round(tenor_years * 12))}M")
        ours = curve.forward_rate(d, d2)
        theirs = ql_curve.forwardRate(
            _to_ql_date(d), _to_ql_date(d2), day_counter, ql.Simple
        ).rate()
        err = abs(ours - theirs)
        max_err = max(max_err, err)
        status = "pass" if err < tolerance else "warn"
        report.items.append(
            ValidationItem(f"fwd.{d.isoformat()}", status, f"|ours - ql| = {err:.3e}", {"ours": ours, "quantlib": theirs})
        )

    report.items.append(
        ValidationItem("fwd.max_error", "pass" if max_err < tolerance else "warn", f"max |ours - ql| = {max_err:.3e}")
    )
    return report
