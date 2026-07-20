"""Ladder L3 — SOFR Index validation (docs/PLAN.md §5.3).

Recompound NY Fed daily SOFR fixings into an index series using the same growth-factor formula
the NY Fed publishes (``openusdcurve.pricing.sofr_compounding.compound_factor_from_fixings``),
anchor it to a published index level, and compare the recomputed series against the published
SOFR Index date-by-date. Returns a per-date error series inside the report's ``items``/details
rather than raising.

Works directly against ``openusdcurve.data.new_york_fed`` output (fixings/index dataframes) or
against a small fixture built with :class:`openusdcurve.pricing.sofr_compounding.Fixing` when no
live data is available (e.g. offline tests).
"""

from __future__ import annotations

from datetime import date

from openusdcurve.pricing.sofr_compounding import Fixing, compound_factor_from_fixings
from openusdcurve.validation.report import ValidationItem, ValidationReport

__all__ = ["sofr_index_recovery_report", "fixings_from_dataframe", "index_from_dataframe"]


def fixings_from_dataframe(df) -> list[Fixing]:
    """Build a ``list[Fixing]`` from a NY Fed-style dataframe with ``effective_date``/``rate``
    columns (see ``openusdcurve.data.new_york_fed._parse_rates_json``)."""
    return [Fixing(effective_date=row.effective_date, rate=float(row.rate)) for row in df.itertuples()]


def index_from_dataframe(df) -> dict[date, float]:
    """Build a ``{date: index_level}`` mapping from a NY Fed-style dataframe with
    ``effective_date``/``index`` columns (see
    ``openusdcurve.data.new_york_fed._parse_index_json``)."""
    return {row.effective_date: float(row.index) for row in df.itertuples()}


def sofr_index_recovery_report(
    fixings: list[Fixing],
    published_index: dict[date, float],
    tolerance: float = 1e-6,
) -> ValidationReport:
    """Recompound daily SOFR fixings and compare to a published index series.

    ``published_index`` maps date -> index level. The earliest date is used as the anchor:
    ``I_model(t) = I_published(t0) * growth(t0, t)``. Each subsequent date contributes one
    per-date error item; a rolled-up ``index.max_rel_error`` item summarizes the series.
    """
    report = ValidationReport(label="L3 SOFR index recovery")

    if not fixings:
        report.items.append(ValidationItem("inputs.fixings", "fail", "no fixings supplied"))
        return report
    if not published_index:
        report.items.append(ValidationItem("inputs.published_index", "fail", "no published index supplied"))
        return report

    dates_sorted = sorted(published_index.keys())
    t0 = dates_sorted[0]
    i0 = published_index[t0]

    max_rel_err = 0.0
    for d in dates_sorted[1:]:
        try:
            factor = compound_factor_from_fixings(fixings, t0, d)
        except ValueError as exc:
            report.items.append(ValidationItem(f"index.{d.isoformat()}", "fail", str(exc)))
            continue

        model_index = i0 * factor
        pub_index = published_index[d]
        rel_err = abs(model_index - pub_index) / pub_index
        max_rel_err = max(max_rel_err, rel_err)
        status = "pass" if rel_err < tolerance else "fail"
        report.items.append(
            ValidationItem(
                f"index.{d.isoformat()}",
                status,
                f"rel error = {rel_err:.3e}",
                {"model_index": model_index, "published_index": pub_index, "rel_error": rel_err},
            )
        )

    report.items.append(
        ValidationItem(
            "index.max_rel_error",
            "pass" if max_rel_err < tolerance else "fail",
            f"max rel error across series = {max_rel_err:.3e}",
            {"max_rel_error": max_rel_err, "tolerance": tolerance},
        )
    )
    return report
