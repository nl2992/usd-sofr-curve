"""Shared OS-1 synthetic dataset: a known reference discount curve and consistent instruments.

The reference curve is built from a smooth continuously-compounded zero curve
``z(T) = 0.030 + 0.004 * T`` (act/365f time). Instruments are constructed so that ALL of their
cashflow dates (accrual endpoints, coupon dates, futures period endpoints) coincide with curve
pillar nodes. That guarantees the bootstrap can recover ``P*`` at the nodes to machine
precision regardless of the interpolation scheme, since ``implied_quote`` only ever evaluates
the curve at exact node dates.
"""

from __future__ import annotations

import math
from datetime import date

from openusdcurve.curves.base import year_fraction as curve_year_fraction
from openusdcurve.curves.discount import DiscountCurve
from openusdcurve.instruments.conventions import add_tenor
from openusdcurve.instruments.deposit import Deposit
from openusdcurve.instruments.future import Future
from openusdcurve.instruments.ois import OIS
from openusdcurve.instruments.swap_libor import SwapLIBOR

VALUATION_DATE = date(2025, 1, 2)


def reference_zero(t: float) -> float:
    """Smooth continuously-compounded zero rate at act/365f time ``t``."""
    return 0.030 + 0.004 * t


def reference_df(d: date) -> float:
    """Reference discount factor P*(valuation, d)."""
    t = curve_year_fraction(VALUATION_DATE, d)
    return math.exp(-reference_zero(t) * t)


def _node(months: int) -> date:
    return add_tenor(VALUATION_DATE, f"{months}M")


# All the curve node dates the instruments will touch (ascending).
NODE_MONTHS = [6, 9, 12, 18, 24, 36]
NODE_DATES = [_node(m) for m in NODE_MONTHS]


def true_curve() -> DiscountCurve:
    """The known reference curve P*, pinned at exactly the node dates."""
    return DiscountCurve(
        reference_date=VALUATION_DATE,
        pillar_dates=list(NODE_DATES),
        discount_factors=[reference_df(d) for d in NODE_DATES],
    )


def build_synthetic_instruments() -> list:
    """Deposit + Future + LIBOR swaps + OIS whose cashflow dates all land on node dates.

    Target quotes are computed from the reference curve so that a correct bootstrap recovers
    P* exactly at every node.
    """
    curve = true_curve()
    n6, n9, n12, n18, n24, n36 = NODE_DATES

    instruments: list = []

    # --- Deposits (act/360), start at valuation, pillars at 6M and 12M ---------------------
    dep6 = Deposit(
        instrument_id="DEP_6M",
        valuation_date=VALUATION_DATE,
        start_date=VALUATION_DATE,
        maturity_date=n6,
        target_quote=0.0,
    )
    dep6.target_quote = dep6.implied_quote(curve)
    instruments.append(dep6)

    dep12 = Deposit(
        instrument_id="DEP_12M",
        valuation_date=VALUATION_DATE,
        start_date=VALUATION_DATE,
        maturity_date=n12,
        target_quote=0.0,
    )
    dep12.target_quote = dep12.implied_quote(curve)
    instruments.append(dep12)

    # --- Future (act/360), 3M-style period [6M, 9M], price = 100*(1 - fwd) -----------------
    fut = Future(
        instrument_id="FUT_6x9",
        valuation_date=VALUATION_DATE,
        period_start=n6,
        period_end=n9,
        target_quote=0.0,
    )
    fut.target_quote = fut.implied_quote(curve)
    instruments.append(fut)

    # --- LIBOR swaps (semiannual 30/360 fixed), coupon dates are all nodes -----------------
    sw18 = SwapLIBOR(
        instrument_id="SWAP_18M",
        valuation_date=VALUATION_DATE,
        effective_date=VALUATION_DATE,
        maturity_date=n18,
        target_quote=0.0,
        fixed_schedule=[n6, n12, n18],
    )
    sw18.target_quote = sw18.implied_quote(curve)
    instruments.append(sw18)

    sw24 = SwapLIBOR(
        instrument_id="SWAP_24M",
        valuation_date=VALUATION_DATE,
        effective_date=VALUATION_DATE,
        maturity_date=n24,
        target_quote=0.0,
        fixed_schedule=[n6, n12, n18, n24],
    )
    sw24.target_quote = sw24.implied_quote(curve)
    instruments.append(sw24)

    # --- OIS (annual act/360 fixed), coupon dates are nodes --------------------------------
    ois36 = OIS(
        instrument_id="OIS_36M",
        valuation_date=VALUATION_DATE,
        effective_date=VALUATION_DATE,
        maturity_date=n36,
        target_quote=0.0,
        fixed_schedule=[n12, n24, n36],
    )
    ois36.target_quote = ois36.implied_quote(curve)
    instruments.append(ois36)

    return instruments
