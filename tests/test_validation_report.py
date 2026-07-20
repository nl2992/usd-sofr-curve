"""run_validation returns a well-formed report; quantlib_compare is skipped gracefully when
QuantLib is absent; compare_to_benchmark handles a benchmark CSV/dataframe.
"""

from __future__ import annotations

import pandas as pd
import pytest

from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.instruments.conventions import add_tenor
from openusdcurve.validation import ValidationReport, compare_to_benchmark, run_validation
from openusdcurve.validation.quantlib_compare import HAS_QUANTLIB, compare_discount_factors
from tests._synthetic import NODE_DATES, VALUATION_DATE, build_synthetic_instruments


def test_run_validation_returns_well_formed_report():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    report = run_validation(boot, instruments)

    assert isinstance(report, ValidationReport)
    assert report.label == "openusdcurve validation"
    assert len(report.items) > 0
    assert all(item.name.startswith(("L4.", "L5.")) for item in report.items)
    # L4 must have passed for a self-consistent bootstrapped curve.
    l4_items = [item for item in report.items if item.name.startswith("L4.")]
    assert all(item.status == "pass" for item in l4_items)
    # to_text() must not raise and should mention the overall status.
    text = report.to_text()
    assert "status:" in text


def test_quantlib_compare_skips_gracefully_when_absent():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    report = compare_discount_factors(boot, NODE_DATES)

    if HAS_QUANTLIB:
        assert report.skipped is False
        assert report.status in ("pass", "warn", "fail")
    else:
        assert report.skipped is True
        assert report.status == "skipped"
        assert any(item.name == "quantlib" for item in report.items)
    # Either way it must never raise and must return a ValidationReport.
    assert isinstance(report, ValidationReport)


def test_compare_to_benchmark_with_tenor_columns():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    tenors = ["6M", "12M", "24M"]
    rows = []
    for t in tenors:
        d = add_tenor(VALUATION_DATE, t)
        rows.append(
            {
                "tenor": t,
                "zero_rate": boot.zero_rate(d) + 0.001,  # deliberately off by 10bp
                "discount": boot.discount(d),
            }
        )
    benchmark_df = pd.DataFrame(rows)

    report = compare_to_benchmark(boot, benchmark_df)

    assert isinstance(report, ValidationReport)
    assert report.label == "benchmark comparison"
    zero_items = [item for item in report.items if item.name.startswith("zero_rate.")]
    discount_items = [item for item in report.items if item.name.startswith("discount.")]
    assert len(zero_items) == len(tenors)
    assert len(discount_items) == len(tenors)
    for item in zero_items:
        assert item.details["delta_bp"] == pytest.approx(-10.0, abs=0.5)
    for item in discount_items:
        assert item.details["delta"] == pytest.approx(0.0, abs=1e-9)


def test_compare_to_benchmark_empty_dataframe_reports_fail_not_raise():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    report = compare_to_benchmark(boot, pd.DataFrame())
    assert report.status == "fail"
    assert isinstance(report, ValidationReport)
