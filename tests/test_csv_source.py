"""Round-trips data/sample/example_quotes.csv into MarketQuote objects via CSVSource."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from openusdcurve.data.base import InstrumentType, LicenseTag, QuoteType
from openusdcurve.data.csv_source import CSVSource

SAMPLE_PATH = Path(__file__).resolve().parents[1] / "data" / "sample" / "example_quotes.csv"


def test_sample_file_exists():
    assert SAMPLE_PATH.exists()


def test_round_trip_all_rows():
    src = CSVSource(SAMPLE_PATH)
    quotes = src.get_quotes()
    assert len(quotes) == 5

    by_id = {q.instrument_id: q for q in quotes}

    depo = by_id["DEPO_3M"]
    assert depo.instrument_type == InstrumentType.DEPOSIT
    assert depo.quote_type == QuoteType.RATE
    assert depo.quote == pytest.approx(0.0525)
    assert depo.valuation_date == date(2026, 7, 20)
    assert depo.maturity_date == date(2026, 10, 20)
    assert depo.bid == pytest.approx(0.0523)
    assert depo.ask == pytest.approx(0.0527)
    assert depo.license == LicenseTag.MANUAL_USER_SUPPLIED
    assert depo.observed_at == datetime(2026, 7, 20, 9, 0, 0)

    fut = by_id["FUT_SEP26"]
    assert fut.instrument_type == InstrumentType.FUTURE
    assert fut.quote_type == QuoteType.PRICE
    assert fut.quote == pytest.approx(94.62)
    assert fut.delayed is True
    assert fut.bid is None and fut.ask is None

    swap = by_id["SWAP_5Y"]
    assert swap.instrument_type == InstrumentType.LIBOR_SWAP
    assert swap.quote_type == QuoteType.PAR_RATE
    assert swap.quote == pytest.approx(0.0412)

    ois = by_id["OIS_5Y"]
    assert ois.instrument_type == InstrumentType.OIS
    assert ois.quote_type == QuoteType.PAR_RATE

    ust = by_id["UST_10Y"]
    assert ust.instrument_type == InstrumentType.TREASURY_PAR
    assert ust.quote_type == QuoteType.RATE


def test_filter_by_valuation_date():
    src = CSVSource(SAMPLE_PATH)
    quotes = src.get_quotes(valuation_date=date(2026, 7, 20))
    assert len(quotes) == 5
    quotes_other = src.get_quotes(valuation_date=date(1999, 1, 1))
    assert quotes_other == []


def test_missing_required_column_raises(tmp_path):
    bad_csv = tmp_path / "bad.csv"
    bad_csv.write_text("valuation_date,instrument_id\n2026-07-20,FOO\n")
    src = CSVSource(bad_csv)
    with pytest.raises(ValueError, match="missing required column"):
        src.get_quotes()


def test_invalid_instrument_type_raises(tmp_path):
    bad_csv = tmp_path / "bad2.csv"
    bad_csv.write_text(
        "valuation_date,instrument_type,instrument_id,maturity_date,quote,quote_type\n"
        "2026-07-20,not_a_type,X,2027-07-20,0.05,rate\n"
    )
    src = CSVSource(bad_csv)
    with pytest.raises(ValueError, match="invalid instrument_type"):
        src.get_quotes()
