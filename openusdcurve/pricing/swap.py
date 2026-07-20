"""Vanilla swap pricing on a curve: fixed annuity, floating leg, par rate, NPV.

Works for both ``SwapLIBOR`` and ``OIS`` calibration instruments since both expose the same
shape (``effective_date``, ``maturity_date``, ``fixed_schedule``, ``fixed_day_count``,
``target_quote``, ``implied_quote``). Deliberately reuses each instrument's own schedule and
day-count (rather than re-deriving a schedule here) so pricing is consistent with the bootstrap
that calibrated the curve — see ``openusdcurve/instruments/swap_libor.py`` and
``openusdcurve/instruments/ois.py`` for the underlying single-curve telescoping identity:
the floating leg (LIBOR or daily-compounded SOFR) always values to
``notional * (P(effective) - P(maturity))`` under single-curve valuation.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Union

from openusdcurve.curves.base import Curve
from openusdcurve.instruments.conventions import year_fraction
from openusdcurve.instruments.ois import OIS
from openusdcurve.instruments.swap_libor import SwapLIBOR

SwapInstrument = Union[SwapLIBOR, OIS]

__all__ = [
    "SwapValuation",
    "fixed_annuity",
    "float_leg_pv",
    "par_rate",
    "price_swap",
]


@dataclass
class SwapValuation:
    """Result of pricing a swap on a curve."""

    par_rate: float
    annuity: float
    fixed_leg_pv: float
    float_leg_pv: float
    npv: float  # payer-fixed convention unless price_swap(payer_fixed=False)


def fixed_annuity(swap: SwapInstrument, curve: Curve) -> float:
    """Sum of ``tau_k * P(t_k)`` over the swap's own fixed schedule/day-count."""
    annuity = 0.0
    prev = swap.effective_date
    for d in swap.fixed_schedule:
        tau = year_fraction(prev, d, swap.fixed_day_count)
        annuity += tau * curve.discount(d)
        prev = d
    return annuity


def float_leg_pv(swap: SwapInstrument, curve: Curve, notional: float = 1.0) -> float:
    """Single-curve telescoping value of the floating leg: ``P(effective) - P(maturity)``."""
    return notional * (curve.discount(swap.effective_date) - curve.discount(swap.maturity_date))


def par_rate(swap: SwapInstrument, curve: Curve) -> float:
    """The rate that makes NPV == 0 on ``curve`` (delegates to the instrument's own formula)."""
    return swap.implied_quote(curve)


def price_swap(
    swap: SwapInstrument,
    curve: Curve,
    fixed_rate: float | None = None,
    notional: float = 1.0,
    payer_fixed: bool = True,
) -> SwapValuation:
    """Price ``swap`` on ``curve`` at ``fixed_rate`` (defaults to the instrument's own quote).

    ``payer_fixed=True`` (default) values the swap from the fixed-rate payer's perspective:
    ``npv = float_leg_pv - fixed_leg_pv``. At ``fixed_rate == par_rate`` this is (near) zero.
    """
    rate = swap.target_quote if fixed_rate is None else fixed_rate
    annuity = fixed_annuity(swap, curve)
    fixed_pv = notional * rate * annuity
    float_pv = float_leg_pv(swap, curve, notional)
    npv = (float_pv - fixed_pv) if payer_fixed else (fixed_pv - float_pv)
    par = par_rate(swap, curve)
    return SwapValuation(
        par_rate=par,
        annuity=annuity,
        fixed_leg_pv=fixed_pv,
        float_leg_pv=float_pv,
        npv=npv,
    )
