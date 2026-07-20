# openusdcurve

Reproducible USD interest-rate curves built primarily from **public data**, with an explicit,
honest split between what is *mathematically functional* and what is *representative of tradable
market levels*.

Every curve carries two independent labels:

| Axis | Values |
| --- | --- |
| Mathematically functional | Yes / No |
| Representative of market levels | High / Medium / Low |

## Curves produced

| Curve | What it is |
| --- | --- |
| `USD-LIBOR-2002-PUBLIC` | Lehman-era single-curve reconstruction from public historical data (Track A). |
| `USD-SOFR-FUTURES-PUBLIC` | Public-data front-end SOFR curve from fixings + 1M/3M futures (Track C / C1). |
| `USD-SOFR-PROXY-PUBLIC` | Futures + explicitly modelled long end (`Y_Treasury + spread`). Proxy only. |
| `USD-SOFR-OIS-MARKET` | Full modern curve using licensed OIS data (Track B). |

> The open-source curves are an independently reproducible research implementation, **not**
> automatically a tradable institutional curve. Proxy sections are labelled and are not suitable
> for valuation, P&L, risk limits, or execution.

## Install (dev)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev,plot]"
```

## CLI

```bash
openusdcurve data pull   --source new-york-fed --date 2026-07-20 [--offline]
openusdcurve build       --config configs/sofr_futures_public.yaml --date 2026-07-20
openusdcurve build       --config configs/lehman_public_2002.yaml  --date 2002-08-26
openusdcurve validate    --curve USD-SOFR-FUTURES-PUBLIC --date 2026-07-20
openusdcurve compare     --curve USD-SOFR-FUTURES-PUBLIC --benchmark bloomberg_curve.csv
```

`--curve` matches a config's `curve_id` (looked up under `--config-dir`, default `configs/`);
`--source` for `data pull` is one of `new-york-fed`, `fred`, `treasury`. Every subcommand accepts
`--offline` to use a bundled sample instead of a live network call, and `build`/`validate`/
`compare` degrade gracefully (pillar table only, or a clear "layer not available" message) if
`openusdcurve.pricing`/`openusdcurve.validation` are not yet installed. See `configs/*.yaml` for
the curve-config schema and `examples/` for runnable, offline end-to-end scripts.

## Design

See [`docs/PLAN.md`](docs/PLAN.md) for the full design, data inventory, validation ladder, and
build phasing. The core code contracts live in `openusdcurve/data/base.py`,
`openusdcurve/curves/base.py`, and `openusdcurve/instruments/base.py`.

## License

MIT. Public-domain data (NY Fed, FRED, US Treasury) may be redistributed; CME delayed quotes are
display-only and licensed vendor data (Bloomberg) is **never** committed.
