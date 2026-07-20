"""Tests for openusdcurve.data.synthetic.SyntheticSource.

Verifies the generated quote set is internally consistent (each quote round-trips exactly
against the closed-form reference curve it was derived from) and covers all instrument types
needed for dataset OS-1 (deposits, futures, LIBOR swaps, OIS).
"""

from __future__ import annotations

import math
from datetime import date

from openusdcurve.data.base import InstrumentType, QuoteType
from openusdcurve.data.synthetic import SyntheticSource


def _act360(d0, d1):
    return (d1 - d0).days / 360.0


def test_covers_all_instrument_types():
    src = SyntheticSource()
    quotes = src.get_quotes(date(2026, 7, 20))
    types = {q.instrument_type for q in quotes}
    assert types == {
        InstrumentType.DEPOSIT,
        InstrumentType.FUTURE,
        InstrumentType.LIBOR_SWAP,
        InstrumentType.OIS,
    }


def test_deterministic_reproducible():
    v = date(2026, 7, 20)
    q1 = SyntheticSource().get_quotes(v)
    q2 = SyntheticSource().get_quotes(v)
    assert [round(q.quote, 12) for q in q1] == [round(q.quote, 12) for q in q2]
    assert [q.instrument_id for q in q1] == [q.instrument_id for q in q2]


def test_deposit_quotes_recover_reference_discount():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    quotes = [q for q in src.get_quotes(v) if q.instrument_type == InstrumentType.DEPOSIT]
    assert len(quotes) == len(src.deposit_tenors_months)
    for q in quotes:
        assert q.quote_type == QuoteType.RATE
        tau = _act360(v, q.maturity_date)
        p_ref = src.reference_discount(v, q.maturity_date)
        implied_p = 1.0 / (1.0 + q.quote * tau)
        assert math.isclose(implied_p, p_ref, rel_tol=1e-12, abs_tol=1e-12)


def test_future_quotes_recover_reference_forward():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    quotes = [q for q in src.get_quotes(v) if q.instrument_type == InstrumentType.FUTURE]
    assert len(quotes) == src.n_futures
    for q in quotes:
        assert q.quote_type == QuoteType.PRICE
        assert q.start_date is not None
        tau = _act360(q.start_date, q.maturity_date)
        implied_fwd = (100.0 - q.quote) / 100.0
        p_start = src.reference_discount(v, q.start_date)
        p_end = src.reference_discount(v, q.maturity_date)
        ref_fwd = (p_start / p_end - 1.0) / tau
        assert math.isclose(implied_fwd, ref_fwd, rel_tol=1e-10, abs_tol=1e-10)
    # consecutive: maturities strictly increasing, start of contract i+1 == end of contract i
    for prev, nxt in zip(quotes, quotes[1:]):
        assert prev.maturity_date == nxt.start_date
        assert nxt.maturity_date > prev.maturity_date


def test_libor_swap_par_rate_prices_to_zero_npv():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    quotes = [q for q in src.get_quotes(v) if q.instrument_type == InstrumentType.LIBOR_SWAP]
    assert len(quotes) == len(src.libor_swap_tenors_years)
    for q in quotes:
        assert q.quote_type == QuoteType.PAR_RATE
        par, maturity = src._par_rate(v, q.maturity_date, 0.5)
        assert math.isclose(par, q.quote, rel_tol=1e-10, abs_tol=1e-10)
        # par swap: fixed leg PV (par * annuity) == floating leg PV (1 - P(T))
        n_periods = round((maturity - v).days / 365.0 / 0.5)
        n_periods = max(n_periods, 1)
        annuity = 0.0
        from openusdcurve.data.synthetic import _add_business_days_roll, _add_months

        for k in range(1, n_periods + 1):
            d = _add_business_days_roll(_add_months(v, round(0.5 * 12 * k)))
            annuity += 0.5 * src.reference_discount(v, d)
        p_final = src.reference_discount(v, maturity)
        assert math.isclose(q.quote * annuity, 1.0 - p_final, rel_tol=1e-9, abs_tol=1e-9)


def test_ois_par_rate_prices_to_zero_npv():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    quotes = [q for q in src.get_quotes(v) if q.instrument_type == InstrumentType.OIS]
    assert len(quotes) == len(src.ois_tenors_years)
    for q in quotes:
        assert q.quote_type == QuoteType.PAR_RATE
        par, maturity = src._par_rate(v, q.maturity_date, 1.0)
        assert math.isclose(par, q.quote, rel_tol=1e-10, abs_tol=1e-10)


def test_reference_discount_at_valuation_date_is_one():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    assert src.reference_discount(v, v) == 1.0


def test_reference_discount_is_decreasing_and_positive():
    src = SyntheticSource()
    v = date(2026, 7, 20)
    quotes = sorted(src.get_quotes(v), key=lambda q: q.maturity_date)
    dfs = [src.reference_discount(v, q.maturity_date) for q in quotes]
    assert all(0.0 < d <= 1.0 for d in dfs)
    assert all(a >= b for a, b in zip(dfs, dfs[1:]))


def test_all_quotes_tagged_with_source_and_valuation_date():
    src = SyntheticSource(name="synthetic-test")
    v = date(2026, 7, 20)
    quotes = src.get_quotes(v)
    assert len(quotes) > 0
    for q in quotes:
        assert q.valuation_date == v
        assert q.source == "synthetic-test"
