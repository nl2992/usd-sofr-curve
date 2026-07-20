"""Builds concrete calibration :class:`Instrument` objects from normalized :class:`MarketQuote`.

The curve engine never learns which data source produced a quote (docs/PLAN.md §4); this
factory is the single place that maps ``InstrumentType`` -> concrete instrument class.
"""

from __future__ import annotations

from datetime import date

from openusdcurve.data.base import InstrumentType, MarketQuote
from openusdcurve.instruments.base import Instrument
from openusdcurve.instruments.conventions import shift_months
from openusdcurve.instruments.deposit import Deposit
from openusdcurve.instruments.future import Future
from openusdcurve.instruments.ois import OIS
from openusdcurve.instruments.swap_libor import SwapLIBOR

# Quote types that describe fixings/reference series, not calibration instruments.
_NON_CALIBRATION_TYPES = {
    InstrumentType.FIXING,
    InstrumentType.SOFR_INDEX,
    InstrumentType.TREASURY_PAR,
}


def build_instruments(quotes: list[MarketQuote], valuation_date: date) -> list[Instrument]:
    """Convert normalized market quotes into concrete calibration instruments.

    Quotes with an ``InstrumentType`` that isn't a calibration instrument (fixings, index
    levels, treasury par yields) are skipped.
    """
    instruments: list[Instrument] = []

    for q in quotes:
        if q.instrument_type in _NON_CALIBRATION_TYPES:
            continue

        start_date = q.start_date or valuation_date

        if q.instrument_type == InstrumentType.DEPOSIT:
            instruments.append(
                Deposit(
                    instrument_id=q.instrument_id,
                    valuation_date=valuation_date,
                    start_date=start_date,
                    maturity_date=q.maturity_date,
                    target_quote=q.quote,
                )
            )
        elif q.instrument_type == InstrumentType.FUTURE:
            period_end = q.maturity_date
            # Futures reference a forward period (e.g. 3M IMM); the data source is expected to
            # supply an explicit start_date. Fall back to a 3M-prior period_start via calendar
            # arithmetic if none was given.
            period_start = q.start_date or shift_months(period_end, -3)
            instruments.append(
                Future(
                    instrument_id=q.instrument_id,
                    valuation_date=valuation_date,
                    period_start=period_start,
                    period_end=period_end,
                    target_quote=q.quote,
                )
            )
        elif q.instrument_type == InstrumentType.LIBOR_SWAP:
            instruments.append(
                SwapLIBOR(
                    instrument_id=q.instrument_id,
                    valuation_date=valuation_date,
                    effective_date=start_date,
                    maturity_date=q.maturity_date,
                    target_quote=q.quote,
                )
            )
        elif q.instrument_type == InstrumentType.OIS:
            instruments.append(
                OIS(
                    instrument_id=q.instrument_id,
                    valuation_date=valuation_date,
                    effective_date=start_date,
                    maturity_date=q.maturity_date,
                    target_quote=q.quote,
                )
            )
        else:  # pragma: no cover - defensive
            raise ValueError(f"Unsupported instrument_type: {q.instrument_type!r}")

    return instruments
