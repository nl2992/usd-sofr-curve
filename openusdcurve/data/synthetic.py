"""Synthetic quote generator (dataset OS-1: exact synthetic recovery).

Generates a self-consistent set of :class:`MarketQuote` objects — deposits, futures, LIBOR
swaps, and OIS — from a smooth, closed-form *reference* zero curve (Nelson-Siegel). Because the
quotes are algebraically derived from the same curve, a correct bootstrap must recover the
reference discount factors to (near) machine precision at every pillar date. This is the
strongest test in the validation ladder (PLAN §5 step 2, §6 OS-1) precisely because the "true"
answer is known in closed form.

The reference curve here is intentionally independent of ``openusdcurve.curves.discount``
(owned by another sub-agent, and not implemented as of this writing) — the data layer must be
able to generate and self-validate quotes without depending on the bootstrap it is testing.

Conventions used to build quotes (documented, not necessarily identical to production market
convention — internal self-consistency is what matters for OS-1):

- Curve time axis: act/365f, continuously-compounded zero rate (matches
  ``openusdcurve.curves.base.year_fraction`` / ``Curve.zero_rate`` semantics).
- Deposits: simple act/360 money-market rate implied by the reference discount factor.
- Futures: consecutive 3-month IMM (3rd-Wednesday) contracts; price = 100 - 100 * F, where F is
  the act/360 simple forward rate between consecutive IMM dates.
- LIBOR swaps: fixed leg semiannual 30/360 (approximated as exactly 0.5y per period), par rate
  solved against the same reference discount curve (single-curve, as in Track A / OS-1).
- OIS: fixed leg annual (1.0y per period) against the same reference discount curve.
"""

from __future__ import annotations

import calendar
import math
from dataclasses import dataclass, field
from datetime import date, timedelta

from openusdcurve.data.base import (
    InstrumentType,
    LicenseTag,
    MarketQuote,
    QuoteType,
)

__all__ = ["NelsonSiegelParams", "SyntheticSource"]


# ---------------------------------------------------------------------------
# Date helpers (kept local/stdlib-only; instruments/conventions.py is owned by
# another sub-agent and may not exist yet).
# ---------------------------------------------------------------------------


def _act360(d0: date, d1: date) -> float:
    return (d1 - d0).days / 360.0


def _act365f(d0: date, d1: date) -> float:
    return (d1 - d0).days / 365.0


def _add_months(d: date, months: int) -> date:
    total = d.month - 1 + months
    year = d.year + total // 12
    month = total % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def _add_business_days_roll(d: date) -> date:
    """Roll forward to the next weekday (Mon-Fri). No holiday calendar (documented simplification)."""
    while d.weekday() >= 5:  # 5=Sat, 6=Sun
        d += timedelta(days=1)
    return d


def _third_wednesday(year: int, month: int) -> date:
    d = date(year, month, 1)
    # weekday(): Mon=0..Sun=6; Wednesday=2
    first_wed_offset = (2 - d.weekday()) % 7
    first_wed = d + timedelta(days=first_wed_offset)
    return first_wed + timedelta(days=14)


def _next_imm_dates(start: date, count: int) -> list[date]:
    """Next ``count`` quarterly (Mar/Jun/Sep/Dec) IMM dates strictly after ``start``."""
    imm_months = (3, 6, 9, 12)
    out: list[date] = []
    year, month = start.year, start.month
    while len(out) < count:
        for m in imm_months:
            if (year, m) < (start.year, start.month):
                continue
            candidate = _third_wednesday(year, m)
            if candidate > start and candidate not in out:
                out.append(candidate)
            if len(out) >= count:
                break
        year += 1
    out.sort()
    return out[:count]


# ---------------------------------------------------------------------------
# Reference curve
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class NelsonSiegelParams:
    """Parameters of the smooth reference zero curve.

    z(t) = beta0 + beta1 * (1 - exp(-t/tau)) / (t/tau)
                 + beta2 * ((1 - exp(-t/tau)) / (t/tau) - exp(-t/tau))

    Defaults produce a plausible, smoothly upward-sloping USD curve (~2% short end rising
    toward ~3.5% long end) with no arbitrage/negative-rate pathologies over 1D-30Y.
    """

    beta0: float = 0.035
    beta1: float = -0.018
    beta2: float = 0.012
    tau: float = 1.8

    def zero_rate(self, t: float) -> float:
        if t <= 0:
            return self.beta0 + self.beta1
        x = t / self.tau
        decay = (1 - math.exp(-x)) / x
        return self.beta0 + self.beta1 * decay + self.beta2 * (decay - math.exp(-x))


@dataclass
class SyntheticSource:
    """Generates a deterministic, internally-consistent synthetic quote set.

    Parameters are all constructor-injectable so callers can reproduce or vary the scenario
    while keeping everything else fixed (dataset OS-1 requires exact reproducibility).
    """

    params: NelsonSiegelParams = field(default_factory=NelsonSiegelParams)
    deposit_tenors_months: tuple[int, ...] = (0, 1, 2, 3, 6, 9, 12)  # 0 == overnight (1 day)
    n_futures: int = 8
    libor_swap_tenors_years: tuple[int, ...] = (2, 3, 4, 5, 7, 10, 15, 20, 30)
    ois_tenors_years: tuple[int, ...] = (1, 2, 3, 5, 7, 10, 15, 20, 30)
    name: str = "synthetic"
    license: LicenseTag = LicenseTag.UNKNOWN  # not applicable / not a real license (see PLAN §3)

    # -- reference curve API (public, so tests / bootstrap-recovery can compare against it) --

    def reference_zero_rate(self, valuation_date: date, d: date) -> float:
        t = _act365f(valuation_date, d)
        return self.params.zero_rate(t)

    def reference_discount(self, valuation_date: date, d: date) -> float:
        if d == valuation_date:
            return 1.0
        t = _act365f(valuation_date, d)
        z = self.reference_zero_rate(valuation_date, d)
        return math.exp(-z * t)

    # -- quote generation --

    def get_quotes(self, valuation_date: date) -> list[MarketQuote]:
        quotes: list[MarketQuote] = []
        quotes.extend(self._deposit_quotes(valuation_date))
        quotes.extend(self._future_quotes(valuation_date))
        quotes.extend(self._libor_swap_quotes(valuation_date))
        quotes.extend(self._ois_quotes(valuation_date))
        return quotes

    def _deposit_quotes(self, valuation_date: date) -> list[MarketQuote]:
        out = []
        for m in self.deposit_tenors_months:
            if m == 0:
                maturity = _add_business_days_roll(valuation_date + timedelta(days=1))
                tenor_label = "ON"
            else:
                maturity = _add_business_days_roll(_add_months(valuation_date, m))
                tenor_label = f"{m}M"
            tau = _act360(valuation_date, maturity)
            p = self.reference_discount(valuation_date, maturity)
            rate = (1.0 / p - 1.0) / tau
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.DEPOSIT,
                    instrument_id=f"DEPO_{tenor_label}",
                    maturity_date=maturity,
                    quote=rate,
                    quote_type=QuoteType.RATE,
                    source=self.name,
                    license=self.license,
                    start_date=valuation_date,
                )
            )
        return out

    def _future_quotes(self, valuation_date: date) -> list[MarketQuote]:
        imm_dates = [valuation_date, *_next_imm_dates(valuation_date, self.n_futures)]
        out = []
        for i in range(len(imm_dates) - 1):
            start, end = imm_dates[i], imm_dates[i + 1]
            tau = _act360(start, end)
            p_start = self.reference_discount(valuation_date, start)
            p_end = self.reference_discount(valuation_date, end)
            fwd = (p_start / p_end - 1.0) / tau
            price = 100.0 - 100.0 * fwd
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.FUTURE,
                    instrument_id=f"FUT_IMM{i + 1}",
                    maturity_date=end,
                    quote=price,
                    quote_type=QuoteType.PRICE,
                    source=self.name,
                    license=self.license,
                    start_date=start,
                )
            )
        return out

    def _par_rate(
        self, valuation_date: date, maturity_date: date, period_years: float
    ) -> tuple[float, float]:
        """Return (par_rate, annuity) for a fixed leg with equal ``period_years`` accrual periods
        from valuation_date to maturity_date, discounted on the reference curve."""
        n_periods = round(_act365f(valuation_date, maturity_date) / period_years)
        n_periods = max(n_periods, 1)
        annuity = 0.0
        dates = []
        for k in range(1, n_periods + 1):
            months = round(period_years * 12 * k)
            d = _add_business_days_roll(_add_months(valuation_date, months))
            dates.append(d)
        for d in dates:
            annuity += period_years * self.reference_discount(valuation_date, d)
        p_final = self.reference_discount(valuation_date, dates[-1])
        par = (1.0 - p_final) / annuity
        return par, dates[-1]

    def _libor_swap_quotes(self, valuation_date: date) -> list[MarketQuote]:
        out = []
        for y in self.libor_swap_tenors_years:
            maturity_target = _add_business_days_roll(_add_months(valuation_date, y * 12))
            par, maturity = self._par_rate(valuation_date, maturity_target, 0.5)
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.LIBOR_SWAP,
                    instrument_id=f"SWAP_{y}Y",
                    maturity_date=maturity,
                    quote=par,
                    quote_type=QuoteType.PAR_RATE,
                    source=self.name,
                    license=self.license,
                    start_date=valuation_date,
                )
            )
        return out

    def _ois_quotes(self, valuation_date: date) -> list[MarketQuote]:
        out = []
        for y in self.ois_tenors_years:
            maturity_target = _add_business_days_roll(_add_months(valuation_date, y * 12))
            par, maturity = self._par_rate(valuation_date, maturity_target, 1.0)
            out.append(
                MarketQuote(
                    valuation_date=valuation_date,
                    instrument_type=InstrumentType.OIS,
                    instrument_id=f"OIS_{y}Y",
                    maturity_date=maturity,
                    quote=par,
                    quote_type=QuoteType.PAR_RATE,
                    source=self.name,
                    license=self.license,
                    start_date=valuation_date,
                )
            )
        return out
