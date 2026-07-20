"""Sequential 1-D bootstrap: builds a :class:`DiscountCurve` from calibration instruments.

Instruments are sorted by ``pillar_date`` and solved one at a time with Brent's method
(``scipy.optimize.brentq``): at each step, the trial discount factor at the new pillar is the
root of ``implied_quote(curve) - target_quote == 0``, holding all previously-solved pillars
fixed. Because the default interpolator (log-linear on discount factors) only affects
extrapolation beyond the last solved pillar when a new point is appended, earlier pillars are
unaffected by later solves.
"""

from __future__ import annotations

from datetime import date

from scipy.optimize import brentq

from openusdcurve.curves.base import Interpolator
from openusdcurve.curves.discount import DiscountCurve
from openusdcurve.instruments.base import Instrument

_DF_LO = 1e-8
_DF_HI = 10.0


def bootstrap(
    instruments: list[Instrument],
    valuation_date: date,
    interpolator: Interpolator | None = None,
) -> DiscountCurve:
    """Bootstrap a discount curve from calibration instruments.

    Instruments sharing the same ``pillar_date`` are solved in the given (stable) order, each
    appending its own pillar node — later instruments at the same date will simply add another
    node at the same x-coordinate, which is fine as long as they are consistent.
    """
    ordered = sorted(instruments, key=lambda inst: inst.pillar_date)

    pillar_dates: list[date] = []
    discount_factors: list[float] = []

    for inst in ordered:
        if inst.pillar_date <= valuation_date:
            raise ValueError(
                f"Instrument {inst.instrument_id!r} has pillar_date <= valuation_date"
            )

        def objective(trial_df: float, _inst: Instrument = inst) -> float:
            trial_curve = DiscountCurve(
                reference_date=valuation_date,
                pillar_dates=[*pillar_dates, _inst.pillar_date],
                discount_factors=[*discount_factors, trial_df],
                interpolator=interpolator,
            )
            return _inst.implied_quote(trial_curve) - _inst.target_quote

        lo, hi = _DF_LO, _DF_HI
        f_lo, f_hi = objective(lo), objective(hi)
        if f_lo == 0.0:
            solved_df = lo
        elif f_hi == 0.0:
            solved_df = hi
        elif f_lo * f_hi > 0:
            raise RuntimeError(
                f"Bootstrap failed to bracket a root for instrument {inst.instrument_id!r} "
                f"(pillar_date={inst.pillar_date}); f(lo)={f_lo}, f(hi)={f_hi}"
            )
        else:
            solved_df = brentq(objective, lo, hi, xtol=1e-14, rtol=1e-14, maxiter=200)

        pillar_dates.append(inst.pillar_date)
        discount_factors.append(solved_df)

    return DiscountCurve(
        reference_date=valuation_date,
        pillar_dates=pillar_dates,
        discount_factors=discount_factors,
        interpolator=interpolator,
    )
