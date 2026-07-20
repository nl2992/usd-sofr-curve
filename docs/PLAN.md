# openusdcurve — Design & Build Plan

A reproducible USD interest-rate curve toolkit built primarily from **public data**, with an
explicit, honest distinction between what is *mathematically functional* and what is
*representative of tradable market levels*.

## 0. Guiding principle

Every curve the project produces carries **two independent classifications**:

| Axis | Values |
| --- | --- |
| Mathematically functional | Yes / No |
| Representative of market levels | High / Medium / Low |

A curve can be mathematically correct while its long end is a weak market representation
(because public OIS quotes are unavailable or delayed). We never conflate the two.

## 1. The three tracks

```
Track A — Lehman 2002 replication   : historical LIBOR deposits + Eurodollar futures + LIBOR swaps → single-curve bootstrap
Track B — Modern market curve       : SOFR fixing + SOFR futures + licensed SOFR OIS → production SOFR curve
Track C — Fully open-source curve   : public data only; proxy/modelled long end; validated + compared to Bloomberg
```

### Named deliverable curves
- `USD-LIBOR-2002-PUBLIC` — historical single-curve reconstruction (Track A, public inputs).
- `USD-SOFR-FUTURES-PUBLIC` — public-data front-end SOFR curve (Track C / Variant C1).
- `USD-SOFR-PROXY-PUBLIC` — public futures + explicitly modelled long end (Variant C3).
- `USD-SOFR-OIS-MARKET` — full modern curve using licensed OIS data (Track B).

## 2. What "functional" means

The open-source curve is *functional* if it can: ingest public data automatically; generate
valid dated instruments; bootstrap positive discount factors; reproduce its own calibration
instruments; generate zero and forward rates; price a vanilla SOFR swap consistently; handle
historical SOFR fixings; rebuild the curve for prior dates; and pass deterministic unit and
recovery tests. This is **separate** from proving a match to the interdealer USD OIS market.

## 3. Data sources (open-source inventory)

| Source | Use | License tag |
| --- | --- | --- |
| NY Fed Markets Data API | Daily SOFR, percentiles, 30/90/180-day averages, **SOFR Index** (authoritative) | public-domain |
| FRED | Backup daily SOFR; historical USD swap rates (2000–2016); historical LIBOR/H.15 | public-domain |
| CME public delayed | 1M/3M SOFR futures settlement, volume, OI (≥10 min delayed; ToS-restricted redistribution) | public-display-only |
| US Treasury | Par yield curve (2000–present) for comparison & proxy long end | public-domain |
| Manual CSV | User-downloaded settlement/quote files | manual-user-supplied |
| Synthetic | Generated from a known reference curve, for recovery tests | n/a |
| Bloomberg / vendor | Benchmark comparison only; **never committed** | licensed |

**Long-end problem:** public sources give SOFR fixings but not a complete 1Y–50Y OIS par grid.
Hence the three modern variants:
- **C1 futures-only** — SOFR fixing + 1M/3M futures. Strong front end only; do not extrapolate to 30Y.
- **C2 futures + public OIS** — add legally-usable OIS quotes; closest open analogue to Bloomberg.
- **C3 futures + modelled long end** — `K_proxy(T) = Y_Treasury(T) + S_estimated(T)`. Labelled "Proxy". Not for valuation/P&L/risk/execution.

## 4. Core code contract (already scaffolded)

Everything downstream depends on these; do not change their shapes without updating this doc.

- `openusdcurve/data/base.py` — `MarketQuote` (normalized quote), `QuoteType`, `InstrumentType`,
  `LicenseTag`, and the `DataSource` protocol. Every provider returns `list[MarketQuote]`.
- `openusdcurve/curves/base.py` — `Curve` interface: `discount(T)`, `zero_rate(T)`,
  `forward_rate(t1, t2)`, plus `DiscountCurve` (log-linear on discount factors) and the
  `Interpolator` protocol.
- `openusdcurve/instruments/base.py` — `Instrument` protocol with
  `implied_quote(curve) -> float` and `pillar_date`. Concrete: `Deposit`, `Future`, `SwapLIBOR`, `OIS`.
- `openusdcurve/curves/bootstrap.py` — `bootstrap(instruments, valuation_date, interp) -> DiscountCurve`
  via sequential 1-D root finding on each pillar.
- `openusdcurve/pricing/` — par-rate / NPV engines for swaps & OIS; SOFR daily compounding.
- `openusdcurve/validation/` — the test ladder (below).

Day count / calendars: implement `act360`, `act365f`, `thirty360` and a US business-day roll in
`openusdcurve/instruments/conventions.py`. Keep them dependency-free.

## 5. Validation ladder

1. **Unit** — each equation in isolation (deposit DF, forward extraction, futures→rate, annuity, floating telescoping, SOFR compounding, interpolation, root finding).
2. **Synthetic recovery (OS-1)** — build quotes from a known `P*(0,T)`, bootstrap, assert `max_T |P_boot − P*| < 1e-10` at nodes.
3. **SOFR Index validation** — recompound NY Fed daily SOFR, compare `I_model(t)` vs `I_NYFed(t)`.
4. **Exact calibration repricing** — every input quote: `eps_i = q_model − q_input` reported in bp + NPV.
5. **QuantLib comparison** — same dates/calendars/day-counts/instruments/interp; differences must be attributable.
6. **Bloomberg comparison** — `Δz(T)`, `ΔK(T)`, decomposed into instrument-set / timestamp / source / convexity / interpolation / convention / proxy error.

## 6. Datasets

- **OS-1** exact synthetic recovery — proves the code (strongest test; true answer known).
- **OS-2** historical public Lehman-era, valuation date **2002-08-26** — reconstructed single LIBOR curve; label "reconstructed using public historical data", **not** "exact Lehman curve".
- **OS-3** current public SOFR — NY Fed fixings + Index + CME settlements + proxy/where-legal OIS.

## 7. CLI surface

```
openusdcurve data pull   --source new-york-fed --date 2026-07-20
openusdcurve build       --config configs/sofr_futures_public.yaml --date 2026-07-20
openusdcurve build       --config configs/lehman_public_2002.yaml  --date 2002-08-26
openusdcurve validate    --curve USD_SOFR_PUBLIC --date 2026-07-20
openusdcurve compare     --curve USD_SOFR_PUBLIC --benchmark bloomberg_curve.csv
```

## 8. Reproducibility

Persist `data/raw/<source>/<retrieval_ts>/`, `data/normalized/<valuation_date>/`,
`data/curves/<curve_id>/<valuation_date>/`, and a SHA-256 hash of every raw input.
Do not commit restricted market data.

## 9. Data-quality controls

Completeness (required tenors, consecutive futures, valid maturities, fixings through date);
freshness (age, settlement/delayed/real-time, quote date == valuation date); units
(`5.25` vs `0.0525`; `95.25` price vs `4.75%`; bp vs pct); outliers (implausible bounds,
crossed bid/ask, duplicates, stale unchanged, missing fixings); licensing tag on every quote.

## 10. Release 0.1 scope (MVP)

NY Fed SOFR ingestion · FRED backup · CME manual-CSV import · Treasury ingestion · synthetic
quote generator · SOFR futures bootstrap · SOFR OIS pricing · LIBOR single-curve pricing ·
log-linear discount interpolation · calibration report · SOFR Index validation · QuantLib
comparison. **Web scraping is NOT a prerequisite** — manual CSV first.

## 11. Build phasing (who does what)

- **Foundation (Opus, done in scaffold):** core contracts in `data/base.py`, `curves/base.py`,
  `instruments/base.py`, `instruments/conventions.py`; empty package `__init__`s; configs; tests skeleton.
- **Sub-agent A — data layer:** `data/{new_york_fed,fred,cme_public,treasury,csv_source,synthetic}.py` + ingestion/quality controls + tests.
- **Sub-agent B — curve math:** `instruments/*` concretes, `curves/{interpolation,bootstrap,discount}.py` + unit + OS-1 recovery tests.
- **Sub-agent C — pricing & validation:** `pricing/*`, `validation/*` (ladder levels 1–5), SOFR Index validation + tests.
- **Sub-agent D — CLI, configs, examples, docs:** `cli.py`, `configs/*.yaml`, `examples/*`, README wiring, end-to-end smoke test.

All agents code against Section 4 contracts and must not alter them.
