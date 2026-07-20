"""Ladder L4 — exact calibration repricing (docs/PLAN.md §5.4).

``reprice_calibration`` is the function the CLI's ``validate`` command calls directly — its
signature is STABLE: ``reprice_calibration(instruments, curve) -> ValidationReport``. Do not
change the parameter order/names without updating the CLI and this docstring together.
"""

from __future__ import annotations

from dataclasses import dataclass

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.base import Instrument
from openusdcurve.validation.report import ValidationItem, ValidationReport

__all__ = ["RepricingResult", "reprice_calibration"]


@dataclass
class RepricingResult:
    """Per-instrument repricing outcome (also stashed in ``ValidationItem.details['result']``)."""

    instrument_id: str
    input_quote: float
    model_quote: float
    error_bp: float
    npv_error: float


def reprice_calibration(
    instruments: list[Instrument],
    curve: Curve,
    tolerance_bp: float = 0.01,
) -> ValidationReport:
    """Reprice every calibration instrument on ``curve`` and report input vs. model quote error.

    For each instrument: ``error = model_quote - input_quote`` (in the instrument's own quote
    units — decimal rate for Deposit/SwapLIBOR/OIS, price for Future), reported in bp
    (``error * 10000``) plus as an ``npv_error`` (same value; a rate/price error IS the NPV
    sensitivity direction for a par instrument). Never raises: a broken ``implied_quote`` call is
    reported as a failing item for that instrument, not propagated.
    """
    report = ValidationReport(label="L4 exact calibration repricing")

    max_err_bp = 0.0
    for inst in instruments:
        try:
            model_quote = inst.implied_quote(curve)
        except Exception as exc:  # pragma: no cover - defensive; report, don't raise
            report.items.append(
                ValidationItem(inst.instrument_id, "fail", f"implied_quote raised: {exc!r}")
            )
            continue

        error = model_quote - inst.target_quote
        error_bp = error * 10000.0
        npv_error = error
        max_err_bp = max(max_err_bp, abs(error_bp))
        status = "pass" if abs(error_bp) < tolerance_bp else "fail"

        result = RepricingResult(
            instrument_id=inst.instrument_id,
            input_quote=inst.target_quote,
            model_quote=model_quote,
            error_bp=error_bp,
            npv_error=npv_error,
        )
        report.items.append(
            ValidationItem(
                inst.instrument_id,
                status,
                f"input={inst.target_quote!r} model={model_quote!r} error={error_bp:.4f} bp",
                {
                    "input": inst.target_quote,
                    "model": model_quote,
                    "error_bp": error_bp,
                    "npv_error": npv_error,
                    "result": result,
                },
            )
        )

    report.items.append(
        ValidationItem(
            "repricing.max_error_bp",
            "pass" if max_err_bp < tolerance_bp else "fail",
            f"max |error| across instruments = {max_err_bp:.4f} bp",
            {"max_error_bp": max_err_bp, "tolerance_bp": tolerance_bp},
        )
    )
    return report
