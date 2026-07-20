"""``openusdcurve`` command-line interface (docs/PLAN.md §7).

    openusdcurve data pull   --source new-york-fed --date 2026-07-20
    openusdcurve build       --config configs/sofr_futures_public.yaml --date 2026-07-20
    openusdcurve validate    --curve USD-SOFR-FUTURES-PUBLIC --date 2026-07-20
    openusdcurve compare     --curve USD-SOFR-FUTURES-PUBLIC --benchmark bloomberg_curve.csv

Every subcommand is designed to degrade gracefully:

- ``data pull`` falls back to a bundled sample payload if a live fetch fails or ``--offline`` is
  passed, and never raises on a network error.
- ``build`` imports ``openusdcurve.pricing`` lazily; if it is not yet installed/built, it prints
  the pillar table only (no par-rate/NPV metrics) rather than failing.
- ``validate``/``compare`` import ``openusdcurve.validation`` lazily; if that package is not yet
  available, they print a clear message and exit 0 (not an error — the layer simply isn't there
  yet), per the task spec.

Exit codes: 0 success (including "layer not available"), 1 usage/expected failure (e.g. missing
config/curve/file), 2 unexpected internal error.
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

__all__ = ["main", "build"]

_DEFAULT_CURVES_DIR = Path("data/curves")

_SOURCE_CHOICES = ("new-york-fed", "fred", "treasury")
_SOURCE_TYPE_MAP = {
    "new-york-fed": "new_york_fed",
    "fred": "fred",
    "treasury": "treasury",
}


# ---------------------------------------------------------------------------
# argparse wiring
# ---------------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="openusdcurve", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_pull = sub.add_parser("data", help="Data-layer operations")
    pull_sub = p_pull.add_subparsers(dest="data_command", required=True)
    p_data_pull = pull_sub.add_parser("pull", help="Fetch (or load a sample of) raw market data")
    p_data_pull.add_argument("--source", required=True, choices=_SOURCE_CHOICES)
    p_data_pull.add_argument("--date", required=True, help="YYYY-MM-DD")
    p_data_pull.add_argument(
        "--offline", action="store_true", help="Skip the network call and use a bundled sample"
    )
    p_data_pull.add_argument("--raw-dir", default="data/raw")
    p_data_pull.add_argument("--normalized-dir", default="data/normalized")
    p_data_pull.set_defaults(func=cmd_data_pull)

    p_build = sub.add_parser("build", help="Build (bootstrap) a curve from a config")
    p_build.add_argument("--config", required=True, help="Path to a configs/*.yaml file")
    p_build.add_argument("--date", default=None, help="YYYY-MM-DD (defaults to config's date)")
    p_build.add_argument(
        "--offline", action="store_true", help="Use bundled samples for any network-backed source"
    )
    p_build.add_argument("--out-dir", default=str(_DEFAULT_CURVES_DIR))
    p_build.add_argument("--no-save", action="store_true", help="Do not persist curve output")
    p_build.set_defaults(func=cmd_build)

    p_validate = sub.add_parser("validate", help="Rebuild a curve and run the validation ladder")
    p_validate.add_argument("--curve", required=True, help="curve_id, e.g. USD-SOFR-FUTURES-PUBLIC")
    p_validate.add_argument("--date", required=True, help="YYYY-MM-DD")
    p_validate.add_argument("--config-dir", default="configs")
    p_validate.add_argument("--offline", action="store_true")
    p_validate.set_defaults(func=cmd_validate)

    p_compare = sub.add_parser("compare", help="Compare a rebuilt curve to a benchmark CSV")
    p_compare.add_argument("--curve", required=True, help="curve_id, e.g. USD-SOFR-FUTURES-PUBLIC")
    p_compare.add_argument("--benchmark", required=True, help="Path to a benchmark CSV")
    p_compare.add_argument("--date", default=None, help="YYYY-MM-DD (defaults to config's date)")
    p_compare.add_argument("--config-dir", default="configs")
    p_compare.add_argument("--offline", action="store_true")
    p_compare.set_defaults(func=cmd_compare)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args) or 0)
    except _CLIError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - defensive top-level guard
        print(f"unexpected error: {exc!r}", file=sys.stderr)
        return 2


class _CLIError(RuntimeError):
    """Expected, user-facing failure (bad config/curve/file) -> exit code 1."""


# ---------------------------------------------------------------------------
# config discovery helper (used by validate/compare, which take --curve not --config)
# ---------------------------------------------------------------------------


def _find_config_for_curve(curve_id: str, config_dir: str):
    from openusdcurve.config import load_curve_config

    config_dir_path = Path(config_dir)
    if not config_dir_path.is_dir():
        raise _CLIError(f"config directory not found: {config_dir_path}")

    for path in sorted(config_dir_path.glob("*.yaml")):
        try:
            cfg = load_curve_config(path)
        except Exception:
            continue
        if cfg.curve_id == curve_id:
            return cfg, path
    raise _CLIError(
        f"no config in {config_dir_path} declares curve_id={curve_id!r} "
        f"(looked at *.yaml files there)"
    )


def _resolve_valuation_date(date_str: str | None, cfg) -> date:
    if date_str is not None:
        return date.fromisoformat(date_str)
    if getattr(cfg, "default_valuation_date", None) is not None:
        return cfg.default_valuation_date
    raise _CLIError("--date is required (config has no default_valuation_date)")


# ---------------------------------------------------------------------------
# data pull
# ---------------------------------------------------------------------------


def cmd_data_pull(args: argparse.Namespace) -> int:
    from openusdcurve.config import _offline_client_for
    from openusdcurve.data.reproducibility import save_normalized, save_raw

    valuation_date = date.fromisoformat(args.date)
    source_type = _SOURCE_TYPE_MAP[args.source]

    used_sample = args.offline
    quotes = []
    raw_note = "no raw payload captured (adapter fetches internally)"

    def _make_source(offline: bool):
        if source_type == "new_york_fed":
            from openusdcurve.data.new_york_fed import NewYorkFedSOFRSource

            client = _offline_client_for(source_type) if offline else None
            return NewYorkFedSOFRSource(client=client)
        if source_type == "fred":
            from openusdcurve.data.fred import FREDSource

            client = _offline_client_for(source_type) if offline else None
            return FREDSource(client=client)
        from openusdcurve.data.treasury import TreasurySource

        client = _offline_client_for(source_type) if offline else None
        return TreasurySource(client=client)

    if not args.offline:
        try:
            source = _make_source(offline=False)
            quotes = source.get_quotes(valuation_date)
        except Exception as exc:
            print(f"live fetch for {args.source!r} failed ({exc!r}); falling back to sample", file=sys.stderr)
            quotes = []
            used_sample = True

    if not quotes:
        used_sample = True
        source = _make_source(offline=True)
        quotes = source.get_quotes(valuation_date)

    raw_meta = save_raw(
        args.source,
        raw="\n".join(repr(q) for q in quotes) or "(no quotes)",
        base_dir=args.raw_dir,
        filename=f"{valuation_date.isoformat()}.txt",
        extra_metadata={"used_sample": used_sample, "valuation_date": valuation_date.isoformat()},
    )
    norm_meta = save_normalized(quotes, valuation_date, base_dir=args.normalized_dir)

    print(f"source:            {args.source}")
    print(f"valuation_date:    {valuation_date}")
    print(f"used bundled sample: {used_sample}")
    print(f"quotes fetched:    {len(quotes)}")
    for q in quotes:
        print(f"  - {q.instrument_type.value:14s} {q.instrument_id:16s} {q.quote_type.value:9s} {q.quote}")
    print(f"raw saved:         {raw_meta['path']} (sha256={raw_meta['sha256'][:12]}...)")
    print(f"normalized saved:  {norm_meta['path']} (sha256={norm_meta['sha256'][:12]}...)")
    return 0


# ---------------------------------------------------------------------------
# build
# ---------------------------------------------------------------------------


def cmd_build(args: argparse.Namespace) -> int:
    from openusdcurve.config import build_curve_from_config, load_curve_config

    config_path = Path(args.config)
    if not config_path.is_file():
        raise _CLIError(f"config file not found: {config_path}")

    cfg = load_curve_config(config_path)
    valuation_date = _resolve_valuation_date(args.date, cfg)

    try:
        curve, quotes, instruments = build_curve_from_config(
            cfg, valuation_date, offline=args.offline
        )
    except FileNotFoundError as exc:
        raise _CLIError(
            f"curve {cfg.curve_id!r}: required data file not found ({exc}). "
            f"See the config's classification.notes for what to supply."
        ) from exc

    print(f"curve_id:          {cfg.curve_id}")
    print(f"description:       {cfg.description.strip()}")
    print(
        f"classification:    functional={cfg.classification.functional} "
        f"market_representativeness={cfg.classification.market_representativeness}"
    )
    print(f"valuation_date:    {valuation_date}")
    print(f"quotes used:       {len(quotes)}")
    print(f"instruments built: {len(instruments)}")
    print()
    print(f"{'pillar_date':12s} {'instrument_id':16s} {'discount_factor':>16s} {'zero_rate':>12s}")
    for inst in sorted(instruments, key=lambda i: i.pillar_date):
        df = curve.discount(inst.pillar_date)
        z = curve.zero_rate(inst.pillar_date)
        print(f"{inst.pillar_date.isoformat():12s} {inst.instrument_id:16s} {df:16.10f} {z:12.6%}")

    _print_pricing_metrics(curve, instruments)

    if not args.no_save:
        out_dir = _save_curve(cfg.curve_id, valuation_date, curve, instruments, args.out_dir)
        print(f"\nsaved curve to:    {out_dir}")

    return 0


def _print_pricing_metrics(curve, instruments) -> None:
    try:
        from openusdcurve.pricing.metrics import par_swap_rates
    except ImportError:
        print("\n(pricing layer not available yet — printed pillar discount/zero rates only)")
        return

    # Report par OIS rates only at standard tenors that fall WITHIN the calibrated span —
    # extrapolated long-tenor par rates off a short curve are meaningless, so we omit them.
    from openusdcurve.instruments.conventions import add_tenor

    candidate_tenors = ["3M", "6M", "1Y", "2Y", "3Y", "5Y", "7Y", "10Y", "20Y", "30Y"]
    last_pillar = max(inst.pillar_date for inst in instruments)
    in_span = [t for t in candidate_tenors if add_tenor(curve.reference_date, t) <= last_pillar]
    if not in_span:
        print(f"\n(curve spans only to {last_pillar.isoformat()}; no standard par tenor within span)")
        return

    try:
        rates = par_swap_rates(curve, in_span)
    except Exception as exc:  # pragma: no cover - defensive; pricing layer may still be WIP
        print(f"\n(pricing metrics unavailable: {exc!r})")
        return

    print("\npar OIS rates (within calibrated span):")
    for tenor in in_span:
        print(f"  {tenor:>4s}  {rates[tenor]:12.6%}")


def _save_curve(curve_id: str, valuation_date: date, curve, instruments, out_dir: str) -> Path:
    import json

    out_path = Path(out_dir) / curve_id / valuation_date.isoformat()
    out_path.mkdir(parents=True, exist_ok=True)

    payload = {
        "curve_id": curve_id,
        "valuation_date": valuation_date.isoformat(),
        "pillars": [
            {
                "instrument_id": inst.instrument_id,
                "pillar_date": inst.pillar_date.isoformat(),
                "discount_factor": curve.discount(inst.pillar_date),
                "zero_rate": curve.zero_rate(inst.pillar_date),
            }
            for inst in sorted(instruments, key=lambda i: i.pillar_date)
        ],
    }
    (out_path / "curve.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return out_path


# ---------------------------------------------------------------------------
# validate
# ---------------------------------------------------------------------------


def cmd_validate(args: argparse.Namespace) -> int:
    from openusdcurve.config import build_curve_from_config

    cfg, _path = _find_config_for_curve(args.curve, args.config_dir)
    valuation_date = _resolve_valuation_date(args.date, cfg)
    curve, _quotes, instruments = build_curve_from_config(cfg, valuation_date, offline=args.offline)

    try:
        from openusdcurve.validation import run_validation
    except ImportError:
        print(
            f"validation layer not available yet — curve {cfg.curve_id!r} was rebuilt "
            f"successfully for {valuation_date}, but no validation ladder could be run."
        )
        return 0

    report = run_validation(curve, instruments)
    print(report.to_text())
    return 0


# ---------------------------------------------------------------------------
# compare
# ---------------------------------------------------------------------------


def cmd_compare(args: argparse.Namespace) -> int:
    import pandas as pd

    from openusdcurve.config import build_curve_from_config

    benchmark_path = Path(args.benchmark)
    if not benchmark_path.is_file():
        raise _CLIError(f"benchmark file not found: {benchmark_path}")

    cfg, _path = _find_config_for_curve(args.curve, args.config_dir)
    valuation_date = _resolve_valuation_date(args.date, cfg)
    curve, _quotes, _instruments = build_curve_from_config(cfg, valuation_date, offline=args.offline)

    try:
        from openusdcurve.validation import compare_to_benchmark
    except ImportError:
        print(
            f"validation layer not available yet — curve {cfg.curve_id!r} was rebuilt "
            f"successfully for {valuation_date}, but no benchmark comparison could be run."
        )
        return 0

    benchmark_df = pd.read_csv(benchmark_path)
    report = compare_to_benchmark(curve, benchmark_df)
    print(report.to_text() if hasattr(report, "to_text") else report)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
