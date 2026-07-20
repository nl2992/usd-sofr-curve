"""Curve interfaces.

A :class:`Curve` maps time to discount factors and derived rates. The default concrete build is
:class:`DiscountCurve`, which stores discount factors at pillar dates and interpolates
log-linearly on them (i.e. linearly on continuously-compounded zero rate * time). Do not change
these signatures without updating ``docs/PLAN.md`` §4.
"""

from __future__ import annotations

from datetime import date
from typing import Protocol, runtime_checkable

DAYS_PER_YEAR = 365.0


def year_fraction(d0: date, d1: date) -> float:
    """Default act/365f year fraction used for curve time axis (NOT instrument accrual)."""
    return (d1 - d0).days / DAYS_PER_YEAR


@runtime_checkable
class Interpolator(Protocol):
    """Interpolates a scalar y over an ordered x grid. Implementations live in ``interpolation.py``."""

    def __call__(self, x_grid: list[float], y_grid: list[float], x: float) -> float:
        ...


@runtime_checkable
class Curve(Protocol):
    reference_date: date

    def discount(self, d: date) -> float:
        """Discount factor P(reference_date, d). Must be positive and P(ref)=1."""
        ...

    def zero_rate(self, d: date) -> float:
        """Continuously-compounded zero rate to ``d`` (act/365f)."""
        ...

    def forward_rate(self, d1: date, d2: date) -> float:
        """Simple forward rate over [d1, d2] implied by discount factors."""
        ...


class DiscountCurve:
    """Concrete curve: pillar discount factors + an interpolator (default log-linear).

    Implemented by sub-agent B in ``curves/discount.py``. This stub fixes the constructor and
    public surface so the data, pricing, and validation layers can import and type against it
    immediately; the real implementation is spliced in at the bottom of this module.
    """

    reference_date: date

    def __init__(
        self,
        reference_date: date,
        pillar_dates: list[date],
        discount_factors: list[float],
        interpolator: Interpolator | None = None,
    ) -> None:
        raise NotImplementedError("Implemented by sub-agent B (curves/discount.py).")

    def discount(self, d: date) -> float:  # pragma: no cover - stub
        raise NotImplementedError

    def zero_rate(self, d: date) -> float:  # pragma: no cover - stub
        raise NotImplementedError

    def forward_rate(self, d1: date, d2: date) -> float:  # pragma: no cover - stub
        raise NotImplementedError


# --- Real implementation splice ---------------------------------------------------------
# ``curves/discount.py`` imports ``Interpolator``/``year_fraction`` from this module, so an
# eager `from openusdcurve.curves.discount import DiscountCurve` here would circular-import
# whenever ``discount.py`` (rather than this module) is the first one loaded. Instead, drop
# the stub name and resolve ``base.DiscountCurve`` lazily via module ``__getattr__`` (PEP 562):
# by the time it's actually accessed, both modules can finish importing each other cleanly.
del DiscountCurve


def __getattr__(name: str):  # noqa: D103
    if name == "DiscountCurve":
        from openusdcurve.curves.discount import DiscountCurve as _DiscountCurve

        return _DiscountCurve
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
