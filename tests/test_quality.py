"""Tests for openusdcurve.data.quality — each check flags the right defect."""

from __future__ import annotations

from datetime import date, datetime, timedelta

from openusdcurve.data.base import InstrumentType, LicenseTag, MarketQuote, QuoteType
from openusdcurve.data.quality import (
    check_completeness,
    check_freshness,
    check_outliers,
    check_units,
    run_quality_checks,
)

V = date(2026, 7, 20)


def _quote(**overrides) -> MarketQuote:
    base = dict(
        valuation_date=V,
        instrument_type=InstrumentType.DEPOSIT,
        instrument_id="DEPO_3M",
        maturity_date=V + timedelta(days=90),
        quote=0.0525,
        quote_type=QuoteType.RATE,
        source="test",
        license=LicenseTag.UNKNOWN,
    )
    base.update(overrides)
    return MarketQuote(**base)


def test_units_flags_percent_instead_of_decimal():
    good = _quote(instrument_id="GOOD", quote=0.0525)
    bad = _quote(instrument_id="BAD", quote=5.25)  # percent, not decimal
    results = check_units([good, bad])
    rate_check = next(r for r in results if r.name == "units.rate_looks_like_percent")
    assert rate_check.status == "fail"
    assert "BAD" in rate_check.details["instrument_ids"]
    assert "GOOD" not in rate_check.details["instrument_ids"]


def test_units_flags_price_looks_like_rate():
    good_future = _quote(
        instrument_id="FUT_GOOD",
        instrument_type=InstrumentType.FUTURE,
        quote=95.25,
        quote_type=QuoteType.PRICE,
    )
    bad_future = _quote(
        instrument_id="FUT_BAD",
        instrument_type=InstrumentType.FUTURE,
        quote=4.75,  # looks like a rate, not a price
        quote_type=QuoteType.PRICE,
    )
    results = check_units([good_future, bad_future])
    price_check = next(r for r in results if r.name == "units.price_looks_like_rate")
    assert price_check.status == "fail"
    assert "FUT_BAD" in price_check.details["instrument_ids"]


def test_outliers_flags_crossed_bid_ask():
    crossed = _quote(instrument_id="CROSSED", bid=0.053, ask=0.051)
    fine = _quote(instrument_id="FINE", bid=0.051, ask=0.053)
    results = check_outliers([crossed, fine])
    crossed_check = next(r for r in results if r.name == "outliers.crossed_bid_ask")
    assert crossed_check.status == "fail"
    assert "CROSSED" in crossed_check.details["instrument_ids"]
    assert "FINE" not in crossed_check.details["instrument_ids"]


def test_outliers_flags_duplicates():
    q1 = _quote(instrument_id="DUP")
    q2 = _quote(instrument_id="DUP")  # same instrument_id + valuation_date
    results = check_outliers([q1, q2])
    dup_check = next(r for r in results if r.name == "outliers.duplicates")
    assert dup_check.status == "fail"
    assert dup_check.details["duplicates"] == [("DUP", V)]


def test_outliers_flags_stale_series():
    history = [
        MarketQuote(
            valuation_date=V - timedelta(days=2),
            instrument_type=InstrumentType.DEPOSIT,
            instrument_id="STALE",
            maturity_date=V + timedelta(days=88),
            quote=0.0525,
            quote_type=QuoteType.RATE,
            source="test",
        ),
        MarketQuote(
            valuation_date=V - timedelta(days=1),
            instrument_type=InstrumentType.DEPOSIT,
            instrument_id="STALE",
            maturity_date=V + timedelta(days=89),
            quote=0.0525,
            quote_type=QuoteType.RATE,
            source="test",
        ),
    ]
    today = _quote(instrument_id="STALE", quote=0.0525)
    results = check_outliers([today], history=history, stale_repeat_threshold=3)
    stale_check = next(r for r in results if r.name == "outliers.stale_series")
    assert stale_check.status == "warn"
    assert "STALE" in stale_check.details["instrument_ids"]


def test_completeness_flags_missing_required_tenor():
    quotes = [_quote(instrument_id="DEPO_3M", instrument_type=InstrumentType.DEPOSIT)]
    results = check_completeness(
        quotes,
        required_instrument_types=(InstrumentType.DEPOSIT, InstrumentType.FUTURE),
        valuation_date=V,
    )
    req_check = next(r for r in results if r.name == "completeness.required_instrument_types")
    assert req_check.status == "fail"
    assert "future" in req_check.details["missing"]


def test_completeness_flags_bad_maturity():
    bad = _quote(instrument_id="BAD_MAT", maturity_date=V)  # not forward-dated
    results = check_completeness([bad], valuation_date=V)
    mat_check = next(r for r in results if r.name == "completeness.valid_maturities")
    assert mat_check.status == "fail"
    assert "BAD_MAT" in mat_check.details["instrument_ids"]


def test_completeness_flags_futures_gap():
    fut1 = _quote(
        instrument_id="FUT1",
        instrument_type=InstrumentType.FUTURE,
        quote=95.0,
        quote_type=QuoteType.PRICE,
        maturity_date=V + timedelta(days=90),
    )
    fut2 = _quote(
        instrument_id="FUT2",
        instrument_type=InstrumentType.FUTURE,
        quote=95.0,
        quote_type=QuoteType.PRICE,
        maturity_date=V + timedelta(days=400),  # big gap
    )
    results = check_completeness([fut1, fut2], valuation_date=V)
    gap_check = next(r for r in results if r.name == "completeness.consecutive_futures")
    assert gap_check.status == "warn"


def test_freshness_flags_wrong_valuation_date():
    wrong = _quote(instrument_id="WRONG", valuation_date=V - timedelta(days=5))
    results = check_freshness([wrong], V)
    vd_check = next(r for r in results if r.name == "freshness.valuation_date_match")
    assert vd_check.status == "fail"
    assert "WRONG" in vd_check.details["instrument_ids"]


def test_freshness_flags_stale_observed_at():
    stale = _quote(
        instrument_id="STALE_OBS",
        observed_at=datetime.combine(V - timedelta(days=10), datetime.min.time()),
    )
    results = check_freshness([stale], V, max_observed_age_days=3)
    obs_check = next(r for r in results if r.name == "freshness.observed_age")
    assert obs_check.status == "warn"


def test_run_quality_checks_aggregates_overall_status():
    good = _quote(instrument_id="GOOD")
    bad = _quote(instrument_id="BAD", quote=5.25)
    report = run_quality_checks([good, bad], V)
    assert report.status == "fail"
    assert len(report.failing()) >= 1
    assert "Overall: fail" in report.summary()


def test_run_quality_checks_all_pass_on_clean_data():
    good = _quote(instrument_id="GOOD_ONLY", bid=0.0523, ask=0.0527)
    report = run_quality_checks([good], V)
    assert report.status in ("pass", "warn")  # no fails on clean input
    assert not report.failing()
