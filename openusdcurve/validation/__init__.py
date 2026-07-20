"""Validation ladder (docs/PLAN.md §5) + the stable API the CLI depends on.

The CLI's ``validate`` and ``compare`` commands (docs/PLAN.md §7) call exactly two functions from
this package — their signatures are STABLE, do not change them without updating the CLI:

- ``run_validation(curve, instruments) -> ValidationReport`` — runs L4 exact-calibration
  repricing plus L5 QuantLib comparison when available (gracefully skipped otherwise).
- ``compare_to_benchmark(curve, benchmark_df) -> ValidationReport`` — compares the curve's zero
  rates / discount factors against a benchmark dataframe (e.g. loaded from a Bloomberg CSV).

The individual ladder rungs (``recovery``, ``sofr_index``, ``repricing``,
``quantlib_compare``) can also be used directly/standalone.
"""

from __future__ import annotations

from datetime import date, datetime

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.base import Instrument
from openusdcurve.validation.quantlib_compare import HAS_QUANTLIB, compare_discount_factors
from openusdcurve.validation.report import Status, ValidationItem, ValidationReport
from openusdcurve.validation.repricing import RepricingResult, reprice_calibration

__all__ = [
    "Status",
    "ValidationItem",
    "ValidationReport",
    "RepricingResult",
    "reprice_calibration",
    "run_validation",
    "compare_to_benchmark",
    "HAS_QUANTLIB",
]


def run_validation(curve: Curve, instruments: list[Instrument]) -> ValidationReport:
    """Run ladder L4 (exact calibration repricing) plus L5 (QuantLib) when available.

    Returns a single aggregated :class:`ValidationReport`. This is the function the CLI's
    ``validate`` command calls directly (docs/PLAN.md §7): ``openusdcurve validate --curve ...``.
    """
    report = ValidationReport(label="openusdcurve validation")

    l4 = reprice_calibration(instruments, curve)
    report.items.extend(
        ValidationItem(f"L4.{item.name}", item.status, item.message, item.details)
        for item in l4.items
    )

    pillar_dates = getattr(curve, "pillar_dates", None) or sorted(
        {inst.pillar_date for inst in instruments}
    )
    l5 = compare_discount_factors(curve, list(pillar_dates))
    if l5.skipped:
        report.items.append(
            ValidationItem("L5.skipped", "skipped", "QuantLib not installed; L5 comparison skipped")
        )
    else:
        report.items.extend(
            ValidationItem(f"L5.{item.name}", item.status, item.message, item.details)
            for item in l5.items
        )

    return report


def compare_to_benchmark(curve: Curve, benchmark_df) -> ValidationReport:
    """Compare ``curve`` to a benchmark dataframe (docs/PLAN.md §5.6 / §7 ``compare`` command).

    ``benchmark_df`` is expected to carry a ``tenor`` (e.g. ``"5Y"``) or ``date`` column plus any
    of ``zero_rate`` / ``discount``; missing columns are simply skipped rather than erroring. This
    is a descriptive comparison (deltas reported, not pass/fail thresholds) since a benchmark
    mismatch is expected and must be decomposed (instrument-set / timestamp / source /
    convexity / interpolation / convention / proxy error) rather than treated as a bug.
    """
    report = ValidationReport(label="benchmark comparison")

    if benchmark_df is None or len(benchmark_df) == 0:
        report.items.append(ValidationItem("benchmark", "fail", "empty benchmark dataframe"))
        return report

    columns = set(benchmark_df.columns)
    has_date = "date" in columns
    has_tenor = "tenor" in columns
    if not has_date and not has_tenor:
        report.items.append(
            ValidationItem("benchmark.columns", "fail", "benchmark must have a 'date' or 'tenor' column")
        )
        return report

    from openusdcurve.instruments.conventions import add_tenor

    for row in benchmark_df.itertuples(index=False):
        row_dict = row._asdict()
        if has_date:
            d = row_dict["date"]
            if isinstance(d, str):
                d = date.fromisoformat(d)
            elif isinstance(d, datetime):
                d = d.date()
            label = d.isoformat()
        else:
            tenor = row_dict["tenor"]
            d = add_tenor(curve.reference_date, tenor)
            label = tenor

        if "zero_rate" in row_dict and row_dict["zero_rate"] is not None:
            model_zero = curve.zero_rate(d)
            bench_zero = float(row_dict["zero_rate"])
            delta_bp = (model_zero - bench_zero) * 10000.0
            report.items.append(
                ValidationItem(
                    f"zero_rate.{label}",
                    "pass",
                    f"delta = {delta_bp:.2f} bp",
                    {"model": model_zero, "benchmark": bench_zero, "delta_bp": delta_bp},
                )
            )

        if "discount" in row_dict and row_dict["discount"] is not None:
            model_df = curve.discount(d)
            bench_df = float(row_dict["discount"])
            delta = model_df - bench_df
            report.items.append(
                ValidationItem(
                    f"discount.{label}",
                    "pass",
                    f"delta = {delta:.3e}",
                    {"model": model_df, "benchmark": bench_df, "delta": delta},
                )
            )

    return report
