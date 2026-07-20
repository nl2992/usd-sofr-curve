"""Ladder L4: reprice_calibration on a bootstrapped synthetic curve gives errors < 0.01bp
for every instrument, and reports a well-formed, non-raising ValidationReport.
"""

from __future__ import annotations

from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.validation.repricing import reprice_calibration
from tests._synthetic import VALUATION_DATE, build_synthetic_instruments


def test_reprice_calibration_all_instruments_within_tolerance():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    report = reprice_calibration(instruments, boot)

    assert report.label == "L4 exact calibration repricing"
    assert report.status == "pass"

    per_instrument_items = [
        item for item in report.items if item.name != "repricing.max_error_bp"
    ]
    assert len(per_instrument_items) == len(instruments)

    for item in per_instrument_items:
        assert item.status == "pass"
        assert abs(item.details["error_bp"]) < 0.01
        assert abs(item.details["npv_error"]) < 1e-6


def test_reprice_calibration_max_error_summary_item():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    report = reprice_calibration(instruments, boot)

    summary = [item for item in report.items if item.name == "repricing.max_error_bp"]
    assert len(summary) == 1
    assert summary[0].status == "pass"
    assert summary[0].details["max_error_bp"] < 0.01


def test_reprice_calibration_reports_rather_than_raises_on_bad_instrument():
    class BrokenInstrument:
        instrument_id = "BROKEN"
        pillar_date = VALUATION_DATE
        target_quote = 0.05

        def implied_quote(self, curve):
            raise RuntimeError("boom")

    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)

    report = reprice_calibration([*instruments, BrokenInstrument()], boot)

    broken_items = [item for item in report.items if item.name == "BROKEN"]
    assert len(broken_items) == 1
    assert broken_items[0].status == "fail"
    assert "boom" in broken_items[0].message
