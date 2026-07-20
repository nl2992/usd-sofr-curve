"""Curve configuration schema + loader (docs/PLAN.md §7, §13).

A curve config is a small, documented YAML file describing:

- ``curve_id`` / ``description`` — identity.
- ``classification`` — the two independent labels every curve carries (PLAN §0):
  ``functional`` (bool) and ``market_representativeness`` (High/Medium/Low), plus free-text
  ``notes`` (e.g. "reconstructed, not exact Lehman", "proxy long end, not for valuation").
- ``source`` — ``type`` (one of the registered :class:`DataSource` adapters, or the composite
  ``sofr_proxy`` builder) plus a free-form ``params`` dict passed to that source's constructor
  (or, for ``sofr_proxy``, to :func:`_build_proxy_quotes`).
- ``instruments`` — optional ``include_types`` filter (list of ``InstrumentType`` values) applied
  to whatever the source returns, plus free-text ``notes`` documenting which tenors/products the
  curve is meant to use.
- ``interpolation`` — one of ``log_linear_discount`` (default), ``linear``, ``flat_forward``.
- ``day_count`` — documentary only at this layer (instruments own their own day counts per
  docs/PLAN.md §4); recorded so a config fully describes what a curve claims to be.

This module intentionally does NOT import ``openusdcurve.pricing`` or ``openusdcurve.validation``
at module scope (they may not exist yet) — only ``openusdcurve.data.*``, ``openusdcurve.curves.*``,
and ``openusdcurve.instruments.*``, all of which are already built.
"""

from __future__ import annotations

from dataclasses import replace
from datetime import date
from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import BaseModel, ConfigDict, Field

from openusdcurve.data.base import DataSource, InstrumentType, LicenseTag, MarketQuote, QuoteType

__all__ = [
    "ClassificationConfig",
    "SourceConfig",
    "InstrumentsConfig",
    "CurveConfig",
    "load_curve_config",
    "build_source",
    "get_quotes_for_config",
    "build_curve_from_config",
]

_SAMPLE_DIR = Path(__file__).resolve().parent.parent / "data" / "sample"

_INTERPOLATORS = {
    "log_linear_discount": "LogLinearDiscount",
    "linear": "Linear",
    "flat_forward": "FlatForward",
}


class ClassificationConfig(BaseModel):
    """The two-axis honesty label every curve carries (PLAN §0)."""

    model_config = ConfigDict(extra="forbid")

    functional: bool = True
    market_representativeness: Literal["High", "Medium", "Low"] = "Medium"
    notes: str | None = None


class SourceConfig(BaseModel):
    """Which :class:`DataSource` adapter to build, and its constructor kwargs.

    ``type`` is one of: ``synthetic``, ``csv``, ``new_york_fed``, ``fred``, ``treasury``,
    ``cme_public``, or the composite ``sofr_proxy`` (futures + Treasury-plus-spread long end,
    PLAN §3 Variant C3 — always labelled proxy / low market-representativeness).
    """

    model_config = ConfigDict(extra="forbid")

    type: Literal[
        "synthetic",
        "csv",
        "new_york_fed",
        "fred",
        "treasury",
        "cme_public",
        "sofr_proxy",
    ]
    params: dict[str, Any] = Field(default_factory=dict)


class InstrumentsConfig(BaseModel):
    """Documents (and optionally filters) which instrument tenors/products feed the curve."""

    model_config = ConfigDict(extra="forbid")

    include_types: list[str] = Field(default_factory=list)
    notes: str | None = None


class CurveConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    curve_id: str
    description: str
    classification: ClassificationConfig
    source: SourceConfig
    instruments: InstrumentsConfig = Field(default_factory=InstrumentsConfig)
    interpolation: Literal["log_linear_discount", "linear", "flat_forward"] = "log_linear_discount"
    day_count: str = "act365f"
    default_valuation_date: date | None = None


def load_curve_config(path: str | Path) -> CurveConfig:
    """Load and validate a curve config YAML file into a :class:`CurveConfig`."""
    text = Path(path).read_text(encoding="utf-8")
    raw = yaml.safe_load(text) or {}
    return CurveConfig(**raw)


# ---------------------------------------------------------------------------
# Offline plumbing: inject a MockTransport httpx.Client built from bundled sample
# files for source types that would otherwise perform real network I/O.
# ---------------------------------------------------------------------------


def _offline_client_for(source_type: str):
    import json

    import httpx

    if source_type == "new_york_fed":
        rates_payload = json.loads((_SAMPLE_DIR / "nyfed_sofr_sample.json").read_text())
        index_payload = json.loads((_SAMPLE_DIR / "nyfed_sofr_index_sample.json").read_text())

        def handler(request: httpx.Request) -> httpx.Response:
            if "sofrai" in str(request.url):
                return httpx.Response(200, json=index_payload)
            return httpx.Response(200, json=rates_payload)

        from openusdcurve.data.new_york_fed import _BASE_URL as _nyfed_base_url

        return httpx.Client(transport=httpx.MockTransport(handler), base_url=_nyfed_base_url)

    if source_type == "fred":
        text = (_SAMPLE_DIR / "fred_sofr_sample.csv").read_text()

        def handler(request: httpx.Request) -> httpx.Response:
            return httpx.Response(200, text=text)

        from openusdcurve.data.fred import _BASE_URL as _fred_base_url

        return httpx.Client(transport=httpx.MockTransport(handler), base_url=_fred_base_url)

    if source_type == "treasury":
        text = (_SAMPLE_DIR / "treasury_sample.csv").read_text()

        def handler(request: httpx.Request) -> httpx.Response:
            return httpx.Response(200, text=text)

        from openusdcurve.data.treasury import _BASE_URL as _treasury_base_url

        return httpx.Client(transport=httpx.MockTransport(handler), base_url=_treasury_base_url)

    return None


def build_source(cfg: SourceConfig, *, offline: bool = False) -> DataSource:
    """Instantiate the :class:`DataSource` adapter described by ``cfg``.

    ``sofr_proxy`` is a composite handled entirely by :func:`get_quotes_for_config`; calling
    ``build_source`` on it raises, since it has no single underlying adapter.
    """
    t = cfg.type
    params = dict(cfg.params)

    if t == "sofr_proxy":
        raise ValueError("'sofr_proxy' is a composite source; use get_quotes_for_config instead")

    if t == "synthetic":
        from openusdcurve.data.synthetic import NelsonSiegelParams, SyntheticSource

        ns_raw = params.pop("nelson_siegel", None)
        if ns_raw is not None:
            params["params"] = NelsonSiegelParams(**ns_raw)
        return SyntheticSource(**params)

    if "license" in params and isinstance(params["license"], str):
        params["license"] = LicenseTag(params["license"])

    if t == "csv":
        from openusdcurve.data.csv_source import CSVSource

        return CSVSource(**params)

    if t == "cme_public":
        from openusdcurve.data.cme_public import CMEPublicDelayedSource

        return CMEPublicDelayedSource(**params)

    if t in ("new_york_fed", "fred", "treasury"):
        if offline:
            params.setdefault("client", _offline_client_for(t))
        if t == "new_york_fed":
            from openusdcurve.data.new_york_fed import NewYorkFedSOFRSource

            return NewYorkFedSOFRSource(**params)
        if t == "fred":
            from openusdcurve.data.fred import FREDSource

            return FREDSource(**params)
        from openusdcurve.data.treasury import TreasurySource

        return TreasurySource(**params)

    raise ValueError(f"Unknown source type: {t!r}")  # pragma: no cover - guarded by pydantic


def _filter_types(quotes: list[MarketQuote], include_types: list[str]) -> list[MarketQuote]:
    if not include_types:
        return quotes
    allowed = set(include_types)
    return [q for q in quotes if q.instrument_type.value in allowed]


def _build_proxy_quotes(
    cfg: SourceConfig, valuation_date: date, *, offline: bool = False
) -> list[MarketQuote]:
    """Variant C3 composite: futures (+ short deposits/fixings) from one source, plus a long end
    modelled as Treasury par yield + a constant spread, re-tagged as OIS pillars.

    ``K_proxy(T) = Y_Treasury(T) + S_estimated(T)`` (PLAN §3). Proxy quotes are always tagged
    with ``license=UNKNOWN`` and an instrument_id prefix of ``PROXY_`` so downstream reports can
    never mistake them for real OIS quotes.
    """
    params = dict(cfg.params)
    spread_bp = float(params.pop("spread_bp", 20.0))
    futures_source_cfg = SourceConfig(**params.pop("futures_source", {"type": "synthetic"}))
    treasury_source_cfg = SourceConfig(**params.pop("treasury_source", {"type": "treasury"}))

    quotes: list[MarketQuote] = []

    fut_source = build_source(futures_source_cfg, offline=offline)
    fut_quotes = fut_source.get_quotes(valuation_date)
    quotes.extend(
        q
        for q in fut_quotes
        if q.instrument_type in (InstrumentType.FUTURE, InstrumentType.DEPOSIT, InstrumentType.FIXING)
    )

    treas_source = build_source(treasury_source_cfg, offline=offline)
    treas_quotes = treas_source.get_quotes(valuation_date)
    # treasury_source may itself be a `csv` adapter carrying a mixed quote file; only treat
    # TREASURY_PAR rows as the long-end proxy input.
    treas_quotes = [q for q in treas_quotes if q.instrument_type == InstrumentType.TREASURY_PAR]
    for q in treas_quotes:
        proxy_rate = q.quote + spread_bp / 10_000.0
        quotes.append(
            replace(
                q,
                instrument_type=InstrumentType.OIS,
                instrument_id=f"PROXY_{q.instrument_id}",
                quote=proxy_rate,
                quote_type=QuoteType.PAR_RATE,
                source=f"{q.source}+proxy-spread",
                license=LicenseTag.UNKNOWN,
                start_date=q.start_date or valuation_date,
            )
        )

    return quotes


def get_quotes_for_config(
    cfg: CurveConfig, valuation_date: date, *, offline: bool = False
) -> list[MarketQuote]:
    """Fetch (or, offline, load bundled samples for) the quotes a curve config describes."""
    if cfg.source.type == "sofr_proxy":
        quotes = _build_proxy_quotes(cfg.source, valuation_date, offline=offline)
    else:
        source = build_source(cfg.source, offline=offline)
        quotes = source.get_quotes(valuation_date)
    return _filter_types(quotes, cfg.instruments.include_types)


def _interpolator_for(name: str):
    from openusdcurve.curves import interpolation as interp_mod

    cls_name = _INTERPOLATORS[name]
    return getattr(interp_mod, cls_name)()


def build_curve_from_config(
    cfg: CurveConfig, valuation_date: date, *, offline: bool = False
) -> tuple[Any, list[MarketQuote], list[Any]]:
    """End-to-end: fetch quotes -> build calibration instruments -> bootstrap.

    Returns ``(curve, quotes, instruments)`` so callers (CLI, tests, examples) can inspect the
    intermediate quotes/instruments as well as the final :class:`DiscountCurve`.
    """
    from openusdcurve.curves.bootstrap import bootstrap
    from openusdcurve.instruments.factory import build_instruments

    quotes = get_quotes_for_config(cfg, valuation_date, offline=offline)
    instruments = build_instruments(quotes, valuation_date)
    if not instruments:
        raise ValueError(
            f"curve {cfg.curve_id!r}: no calibration instruments were built from the "
            f"configured source for {valuation_date} (check source/params/include_types)"
        )
    interpolator = _interpolator_for(cfg.interpolation)
    curve = bootstrap(instruments, valuation_date, interpolator=interpolator)
    return curve, quotes, instruments
