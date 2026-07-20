"""Smoke tests for configs/*.yaml + openusdcurve/config.py (no pricing/validation dependency)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from openusdcurve.config import (
    CurveConfig,
    build_curve_from_config,
    get_quotes_for_config,
    load_curve_config,
)

_CONFIGS_DIR = Path(__file__).resolve().parent.parent / "configs"

_ALL_CONFIG_PATHS = sorted(_CONFIGS_DIR.glob("*.yaml"))


def test_configs_directory_has_the_four_named_curves():
    curve_ids = {load_curve_config(p).curve_id for p in _ALL_CONFIG_PATHS}
    assert curve_ids == {
        "USD-SOFR-FUTURES-PUBLIC",
        "USD-LIBOR-2002-PUBLIC",
        "USD-SOFR-PROXY-PUBLIC",
        "USD-SOFR-OIS-MARKET",
    }


@pytest.mark.parametrize("path", _ALL_CONFIG_PATHS, ids=lambda p: p.name)
def test_config_loads_and_validates(path):
    cfg = load_curve_config(path)
    assert isinstance(cfg, CurveConfig)
    assert cfg.curve_id
    assert cfg.classification.market_representativeness in {"High", "Medium", "Low"}
    assert cfg.interpolation in {"log_linear_discount", "linear", "flat_forward"}


def test_sofr_futures_public_builds_positive_discount_factors():
    cfg = load_curve_config(_CONFIGS_DIR / "sofr_futures_public.yaml")
    curve, quotes, instruments = build_curve_from_config(cfg, date(2026, 7, 20))

    assert quotes
    assert instruments
    for inst in instruments:
        df = curve.discount(inst.pillar_date)
        assert df > 0.0


def test_lehman_public_2002_builds_positive_discount_factors():
    cfg = load_curve_config(_CONFIGS_DIR / "lehman_public_2002.yaml")
    valuation_date = cfg.default_valuation_date
    assert valuation_date == date(2002, 8, 26)

    curve, quotes, instruments = build_curve_from_config(cfg, valuation_date)
    assert quotes
    assert instruments
    for inst in instruments:
        assert curve.discount(inst.pillar_date) > 0.0


def test_sofr_proxy_public_builds_and_includes_proxy_ois_pillars():
    cfg = load_curve_config(_CONFIGS_DIR / "sofr_proxy_public.yaml")
    curve, quotes, instruments = build_curve_from_config(cfg, date(2026, 7, 20))

    assert quotes
    assert any(q.instrument_id.startswith("PROXY_") for q in quotes)
    for inst in instruments:
        assert curve.discount(inst.pillar_date) > 0.0


def test_sofr_market_requires_uncommitted_licensed_data():
    """USD-SOFR-OIS-MARKET points at a licensed file this repo intentionally does not ship."""
    cfg = load_curve_config(_CONFIGS_DIR / "sofr_market.yaml")
    assert cfg.classification.market_representativeness == "High"
    with pytest.raises(FileNotFoundError):
        get_quotes_for_config(cfg, date(2026, 7, 20))


def test_include_types_filters_quotes():
    cfg = load_curve_config(_CONFIGS_DIR / "sofr_futures_public.yaml")
    quotes = get_quotes_for_config(cfg, date(2026, 7, 20))
    seen_types = {q.instrument_type.value for q in quotes}
    assert seen_types <= {"deposit", "future"}
