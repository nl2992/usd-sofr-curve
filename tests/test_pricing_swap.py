"""A par swap priced on the curve it was calibrated to has ~0 NPV; par rate reproduces input."""

from __future__ import annotations

import pytest

from openusdcurve.curves.bootstrap import bootstrap
from openusdcurve.pricing.swap import fixed_annuity, float_leg_pv, par_rate, price_swap
from tests._synthetic import VALUATION_DATE, build_synthetic_instruments


def _swaps(instruments):
    from openusdcurve.instruments.ois import OIS
    from openusdcurve.instruments.swap_libor import SwapLIBOR

    return [inst for inst in instruments if isinstance(inst, (SwapLIBOR, OIS))]


def test_par_rate_reproduces_input_quote():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    for swap in _swaps(instruments):
        implied = par_rate(swap, boot)
        assert implied == pytest.approx(swap.target_quote, abs=1e-10)


def test_price_swap_at_par_has_zero_npv():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    for swap in _swaps(instruments):
        valuation = price_swap(swap, boot)  # defaults to swap.target_quote (== par)
        assert valuation.npv == pytest.approx(0.0, abs=1e-8)
        assert valuation.par_rate == pytest.approx(swap.target_quote, abs=1e-10)


def test_price_swap_off_market_has_nonzero_npv():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    swap = _swaps(instruments)[0]
    off_market_rate = swap.target_quote + 0.01  # 100bp above par
    valuation = price_swap(swap, boot, fixed_rate=off_market_rate)
    annuity = fixed_annuity(swap, boot)
    expected_npv = float_leg_pv(swap, boot) - off_market_rate * annuity
    assert valuation.npv == pytest.approx(expected_npv, abs=1e-12)
    assert valuation.npv != pytest.approx(0.0, abs=1e-6)


def test_float_leg_pv_matches_telescoping_identity():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    swap = _swaps(instruments)[0]
    expected = boot.discount(swap.effective_date) - boot.discount(swap.maturity_date)
    assert float_leg_pv(swap, boot) == pytest.approx(expected, abs=1e-15)


def test_payer_fixed_flips_sign():
    instruments = build_synthetic_instruments()
    boot = bootstrap(instruments, VALUATION_DATE)
    swap = _swaps(instruments)[0]
    off_market_rate = swap.target_quote + 0.01
    payer = price_swap(swap, boot, fixed_rate=off_market_rate, payer_fixed=True)
    receiver = price_swap(swap, boot, fixed_rate=off_market_rate, payer_fixed=False)
    assert payer.npv == pytest.approx(-receiver.npv, abs=1e-12)
