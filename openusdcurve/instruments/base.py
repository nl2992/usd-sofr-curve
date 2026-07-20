"""Instrument contract for bootstrapping.

Each calibration instrument knows its ``pillar_date`` (the curve node it pins) and can compute
its ``implied_quote(curve)`` from a trial curve. The bootstrap solves, pillar by pillar, for the
discount factor that makes ``implied_quote == target_quote``. Concrete instruments (Deposit,
Future, SwapLIBOR, OIS) are implemented by sub-agent B in this package.
"""

from __future__ import annotations

from datetime import date
from typing import Protocol, runtime_checkable

from openusdcurve.curves.base import Curve


@runtime_checkable
class Instrument(Protocol):
    """A calibration instrument. ``target_quote`` is in the same units the source produced."""

    instrument_id: str
    pillar_date: date
    target_quote: float

    def implied_quote(self, curve: Curve) -> float:
        """Model quote (rate/price/par-rate) given a trial curve, in the instrument's own units."""
        ...
