"""Reporting helpers built on top of a calibrated curve: zero rates, forwards, par swap rates."""

from __future__ import annotations

from datetime import date

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import add_tenor
from openusdcurve.instruments.ois import OIS
from openusdcurve.pricing.swap import par_rate

__all__ = ["zero_rates", "forward_curve", "par_swap_rates"]


def zero_rates(curve: Curve, dates: list[date]) -> dict[date, float]:
    """Continuously-compounded zero rate at each of ``dates``."""
    return {d: curve.zero_rate(d) for d in dates}


def forward_curve(curve: Curve, dates: list[date], tenor: str) -> dict[date, float]:
    """Simple forward rate over ``[d, d + tenor]`` for each ``d`` in ``dates``."""
    out: dict[date, float] = {}
    for d in dates:
        d2 = add_tenor(d, tenor)
        out[d] = curve.forward_rate(d, d2)
    return out


def par_swap_rates(
    curve: Curve,
    tenors: list[str],
    valuation_date: date | None = None,
    fixed_frequency: str = "1Y",
    fixed_day_count: str = "act360",
) -> dict[str, float]:
    """Par (OIS-style, annual act/360 fixed) swap rate for each tenor, off ``curve``.

    Builds a throwaway ``OIS`` instrument per tenor purely to reuse its annuity/par-rate
    formula — no bootstrap or calibration is involved, this is read-only curve reporting.
    """
    val = valuation_date or curve.reference_date
    out: dict[str, float] = {}
    for t in tenors:
        maturity = add_tenor(val, t)
        swap = OIS(
            instrument_id=f"PAR_{t}",
            valuation_date=val,
            effective_date=val,
            maturity_date=maturity,
            target_quote=0.0,
            fixed_frequency=fixed_frequency,
            fixed_day_count=fixed_day_count,
        )
        out[t] = par_rate(swap, curve)
    return out
