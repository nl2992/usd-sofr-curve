"""Data-quality controls (PLAN §9).

Four control families, each returning a list of :class:`CheckResult` (pass/warn/fail):

- **Completeness** — required tenors present, consecutive futures (no gaps in IMM sequence),
  valid (non-degenerate, forward-dated) maturities, fixings present through the valuation date.
- **Freshness** — quote age vs valuation date, delayed-flag consistency, ``quote_date ==
  valuation_date`` where that is expected (e.g. fixings).
- **Units** — catches the classic percent-vs-decimal bug (``5.25`` instead of ``0.0525``) and
  price-vs-rate confusion (a ``PRICE`` quote_type value that looks like a decimal rate, or a
  ``RATE``/``PAR_RATE`` value that looks like a futures price).
- **Outliers** — implausible absolute bounds, crossed bid/ask, duplicate quotes, stale
  (unchanged-for-N-observations) series.

None of these checks raise on warnings — only programming errors (bad input types) raise.
Failing checks are still reported (not raised) so a full report can always be produced; callers
decide whether a ``fail`` blocks a bootstrap.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from typing import Literal

from openusdcurve.data.base import InstrumentType, MarketQuote, QuoteType

__all__ = [
    "CheckStatus",
    "CheckResult",
    "QualityReport",
    "check_completeness",
    "check_freshness",
    "check_units",
    "check_outliers",
    "run_quality_checks",
]

CheckStatus = Literal["pass", "warn", "fail"]

# Sanity bounds for the "units" check: any RATE/PAR_RATE quote outside this band is almost
# certainly a percent-vs-decimal mixup (or a genuinely broken input) rather than a real rate.
_PLAUSIBLE_RATE_BOUNDS = (-0.02, 0.30)  # -2% .. 30% decimal
_PLAUSIBLE_PRICE_BOUNDS = (50.0, 110.0)  # futures price band (100 - rate%, sanity generous)
_PLAUSIBLE_INDEX_BOUNDS = (0.5, 5.0)  # SOFR-Index-like level, generous multi-decade band


@dataclass
class CheckResult:
    name: str
    status: CheckStatus
    message: str
    details: dict = field(default_factory=dict)


@dataclass
class QualityReport:
    checks: list[CheckResult] = field(default_factory=list)

    @property
    def status(self) -> CheckStatus:
        statuses = {c.status for c in self.checks}
        if "fail" in statuses:
            return "fail"
        if "warn" in statuses:
            return "warn"
        return "pass"

    def failing(self) -> list[CheckResult]:
        return [c for c in self.checks if c.status == "fail"]

    def warnings(self) -> list[CheckResult]:
        return [c for c in self.checks if c.status == "warn"]

    def summary(self) -> str:
        lines = [f"Overall: {self.status}"]
        for c in self.checks:
            lines.append(f"  [{c.status.upper():4s}] {c.name}: {c.message}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Completeness
# ---------------------------------------------------------------------------


def check_completeness(
    quotes: list[MarketQuote],
    *,
    required_instrument_types: tuple[InstrumentType, ...] = (),
    valuation_date: date | None = None,
) -> list[CheckResult]:
    results = []

    present_types = {q.instrument_type for q in quotes}
    missing_types = [t for t in required_instrument_types if t not in present_types]
    if missing_types:
        results.append(
            CheckResult(
                "completeness.required_instrument_types",
                "fail",
                f"missing required instrument type(s): {[t.value for t in missing_types]}",
                {"missing": [t.value for t in missing_types]},
            )
        )
    else:
        results.append(
            CheckResult(
                "completeness.required_instrument_types", "pass", "all required types present"
            )
        )

    # Valid maturities: maturity must be strictly after valuation_date (or start_date if given).
    bad_maturities = [
        q
        for q in quotes
        if q.maturity_date <= (q.start_date or q.valuation_date)
    ]
    if bad_maturities:
        results.append(
            CheckResult(
                "completeness.valid_maturities",
                "fail",
                f"{len(bad_maturities)} quote(s) with maturity_date <= start/valuation date",
                {"instrument_ids": [q.instrument_id for q in bad_maturities]},
            )
        )
    else:
        results.append(
            CheckResult("completeness.valid_maturities", "pass", "all maturities are forward-dated")
        )

    # Consecutive futures: no gaps in a sorted-by-maturity FUTURE strip (each contract's
    # maturity should roughly follow the previous one; a large gap suggests a missing contract).
    futures = sorted((q for q in quotes if q.instrument_type == InstrumentType.FUTURE), key=lambda q: q.maturity_date)
    gaps = []
    for prev, nxt in zip(futures, futures[1:], strict=False):
        span_days = (nxt.maturity_date - prev.maturity_date).days
        if span_days > 100:  # quarterly IMM contracts are ~91 days apart
            gaps.append((prev.instrument_id, nxt.instrument_id, span_days))
    if gaps:
        results.append(
            CheckResult(
                "completeness.consecutive_futures",
                "warn",
                f"{len(gaps)} gap(s) in futures strip larger than ~1 quarter",
                {"gaps": gaps},
            )
        )
    elif futures:
        results.append(
            CheckResult("completeness.consecutive_futures", "pass", "futures strip is consecutive")
        )

    # Fixings through date: a FIXING/SOFR_INDEX quote's maturity_date (observation date) should
    # be no earlier than a small lag before valuation_date.
    if valuation_date is not None:
        fixings = [
            q for q in quotes if q.instrument_type in (InstrumentType.FIXING, InstrumentType.SOFR_INDEX)
        ]
        stale_fixings = [q for q in fixings if (valuation_date - q.maturity_date).days > 5]
        if stale_fixings:
            results.append(
                CheckResult(
                    "completeness.fixings_through_date",
                    "warn",
                    f"{len(stale_fixings)} fixing(s) more than 5 days behind valuation_date",
                    {"instrument_ids": [q.instrument_id for q in stale_fixings]},
                )
            )
        elif fixings:
            results.append(
                CheckResult("completeness.fixings_through_date", "pass", "fixings are current")
            )

    return results


# ---------------------------------------------------------------------------
# Freshness
# ---------------------------------------------------------------------------


def check_freshness(
    quotes: list[MarketQuote],
    valuation_date: date,
    *,
    max_observed_age_days: int = 3,
) -> list[CheckResult]:
    results = []

    wrong_valuation_date = [q for q in quotes if q.valuation_date != valuation_date]
    if wrong_valuation_date:
        results.append(
            CheckResult(
                "freshness.valuation_date_match",
                "fail",
                f"{len(wrong_valuation_date)} quote(s) with valuation_date != {valuation_date}",
                {"instrument_ids": [q.instrument_id for q in wrong_valuation_date]},
            )
        )
    else:
        results.append(
            CheckResult("freshness.valuation_date_match", "pass", "all quotes match valuation_date")
        )

    stale_observed = []
    for q in quotes:
        if q.observed_at is None:
            continue
        age_days = (valuation_date - q.observed_at.date()).days
        if age_days > max_observed_age_days:
            stale_observed.append((q.instrument_id, age_days))
    if stale_observed:
        results.append(
            CheckResult(
                "freshness.observed_age",
                "warn",
                f"{len(stale_observed)} quote(s) observed more than {max_observed_age_days}d before valuation_date",
                {"details": stale_observed},
            )
        )
    else:
        results.append(CheckResult("freshness.observed_age", "pass", "all observations are fresh"))

    delayed_undeclared = [
        q for q in quotes if q.instrument_type == InstrumentType.FUTURE and q.delayed is False and q.source.startswith("cme")
    ]
    if delayed_undeclared:
        results.append(
            CheckResult(
                "freshness.delayed_flag_consistency",
                "warn",
                f"{len(delayed_undeclared)} CME-sourced future(s) not flagged delayed=True",
                {"instrument_ids": [q.instrument_id for q in delayed_undeclared]},
            )
        )
    else:
        results.append(
            CheckResult("freshness.delayed_flag_consistency", "pass", "delayed flags look consistent")
        )

    return results


# ---------------------------------------------------------------------------
# Units
# ---------------------------------------------------------------------------


def check_units(quotes: list[MarketQuote]) -> list[CheckResult]:
    results = []

    percent_not_decimal = [
        q
        for q in quotes
        if q.quote_type in (QuoteType.RATE, QuoteType.PAR_RATE)
        and not (_PLAUSIBLE_RATE_BOUNDS[0] <= q.quote <= _PLAUSIBLE_RATE_BOUNDS[1])
    ]
    if percent_not_decimal:
        results.append(
            CheckResult(
                "units.rate_looks_like_percent",
                "fail",
                f"{len(percent_not_decimal)} rate quote(s) outside plausible decimal band "
                f"{_PLAUSIBLE_RATE_BOUNDS} — likely percent-not-decimal (e.g. 5.25 vs 0.0525)",
                {"instrument_ids": [q.instrument_id for q in percent_not_decimal]},
            )
        )
    else:
        results.append(CheckResult("units.rate_looks_like_percent", "pass", "rate quotes are decimal-scaled"))

    price_looks_like_rate = [
        q
        for q in quotes
        if q.quote_type == QuoteType.PRICE
        and not (_PLAUSIBLE_PRICE_BOUNDS[0] <= q.quote <= _PLAUSIBLE_PRICE_BOUNDS[1])
    ]
    if price_looks_like_rate:
        results.append(
            CheckResult(
                "units.price_looks_like_rate",
                "fail",
                f"{len(price_looks_like_rate)} PRICE quote(s) outside plausible futures-price "
                f"band {_PLAUSIBLE_PRICE_BOUNDS} — likely rate-not-price confusion (e.g. 4.75 vs 95.25)",
                {"instrument_ids": [q.instrument_id for q in price_looks_like_rate]},
            )
        )
    else:
        results.append(CheckResult("units.price_looks_like_rate", "pass", "price quotes are price-scaled"))

    index_out_of_band = [
        q
        for q in quotes
        if q.quote_type == QuoteType.INDEX
        and not (_PLAUSIBLE_INDEX_BOUNDS[0] <= q.quote <= _PLAUSIBLE_INDEX_BOUNDS[1])
    ]
    if index_out_of_band:
        results.append(
            CheckResult(
                "units.index_out_of_band",
                "warn",
                f"{len(index_out_of_band)} INDEX quote(s) outside plausible band {_PLAUSIBLE_INDEX_BOUNDS}",
                {"instrument_ids": [q.instrument_id for q in index_out_of_band]},
            )
        )
    else:
        results.append(CheckResult("units.index_out_of_band", "pass", "index quotes in plausible band"))

    return results


# ---------------------------------------------------------------------------
# Outliers
# ---------------------------------------------------------------------------


def check_outliers(
    quotes: list[MarketQuote],
    *,
    history: list[MarketQuote] | None = None,
    stale_repeat_threshold: int = 3,
) -> list[CheckResult]:
    results = []

    crossed = [q for q in quotes if q.bid is not None and q.ask is not None and q.bid > q.ask]
    if crossed:
        results.append(
            CheckResult(
                "outliers.crossed_bid_ask",
                "fail",
                f"{len(crossed)} quote(s) with bid > ask",
                {"instrument_ids": [q.instrument_id for q in crossed]},
            )
        )
    else:
        results.append(CheckResult("outliers.crossed_bid_ask", "pass", "no crossed bid/ask"))

    key_counts = Counter((q.instrument_id, q.valuation_date) for q in quotes)
    dupes = [k for k, c in key_counts.items() if c > 1]
    if dupes:
        results.append(
            CheckResult(
                "outliers.duplicates",
                "fail",
                f"{len(dupes)} duplicate (instrument_id, valuation_date) combination(s)",
                {"duplicates": dupes},
            )
        )
    else:
        results.append(CheckResult("outliers.duplicates", "pass", "no duplicate quotes"))

    implausible = [
        q
        for q in quotes
        if q.quote_type in (QuoteType.RATE, QuoteType.PAR_RATE) and not (-0.05 <= q.quote <= 0.50)
    ]
    if implausible:
        results.append(
            CheckResult(
                "outliers.implausible_bounds",
                "fail",
                f"{len(implausible)} rate quote(s) with implausible absolute value",
                {"instrument_ids": [q.instrument_id for q in implausible]},
            )
        )
    else:
        results.append(CheckResult("outliers.implausible_bounds", "pass", "no implausible rate values"))

    if history:
        by_id: dict[str, list[MarketQuote]] = {}
        for q in list(history) + list(quotes):
            by_id.setdefault(q.instrument_id, []).append(q)
        stale = []
        for instrument_id, series in by_id.items():
            series_sorted = sorted(series, key=lambda q: q.valuation_date)
            if len(series_sorted) < stale_repeat_threshold:
                continue
            last_values = [q.quote for q in series_sorted[-stale_repeat_threshold:]]
            if len(set(last_values)) == 1:
                stale.append(instrument_id)
        if stale:
            results.append(
                CheckResult(
                    "outliers.stale_series",
                    "warn",
                    f"{len(stale)} instrument(s) unchanged for {stale_repeat_threshold}+ consecutive observations",
                    {"instrument_ids": stale},
                )
            )
        else:
            results.append(CheckResult("outliers.stale_series", "pass", "no stale/unchanged series detected"))

    return results


def run_quality_checks(
    quotes: list[MarketQuote],
    valuation_date: date,
    *,
    required_instrument_types: tuple[InstrumentType, ...] = (),
    history: list[MarketQuote] | None = None,
) -> QualityReport:
    """Run all four control families and return an aggregate :class:`QualityReport`."""
    report = QualityReport()
    report.checks.extend(
        check_completeness(
            quotes, required_instrument_types=required_instrument_types, valuation_date=valuation_date
        )
    )
    report.checks.extend(check_freshness(quotes, valuation_date))
    report.checks.extend(check_units(quotes))
    report.checks.extend(check_outliers(quotes, history=history))
    return report
