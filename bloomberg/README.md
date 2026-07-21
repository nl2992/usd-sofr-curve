# Bloomberg data template — `USD_SOFR_Curve_Bloomberg.xlsx`

A live Bloomberg workbook for pulling every input needed to bootstrap the **modern** USD SOFR
curve — the SOFR-era replication of the Lehman (2002) single-curve build. It ships **formulas
only, no data**: open it in Excel with the Bloomberg Add-in (BLPAPI) connected and the
`BDP` / `BDH` / `BDS` cells populate. Off-terminal, Bloomberg cells show `#NAME?`/blanks — that
is expected.

## Tabs

| Tab | Contents |
| --- | --- |
| **Instructions** | Parameters (valuation date, history window, price source, tickers as named ranges), the Lehman→modern mapping, and the "do not substitute" warnings (Term SOFR, ICE Swap Rate). |
| **SOFR_Fixings** | Overnight SOFR (`SOFRRATE Index`): current `BDP` + daily `BDH` history. |
| **SOFR_Futures** | SR3 (`SFRA Comdty`, 3M ×20) and SR1 (`SERA Comdty`, 1M ×13) chains via `BDS(...,"FUT_CHAIN")` that spill contract tickers down column A; per-contract fields + implied rate `= 100 − PX_LAST`. |
| **SOFR_OIS_Quotes** | Full `USOSFR…` OIS strip 1W–50Y: bid/ask/last + mid `=(bid+ask)/2` (falls back to last); recommended bootstrap set flagged. **Maturity date and Tenor (yrs) are live formulas parsed from the tenor label** (e.g. `"18M"` → `EDATE(spot,18)`) off the T+2 spot, so they update with the valuation date. |
| **Bootstrap** | **Purely-OIS** SOFR curve: short single-payment OIS (≤1Y) + annual OIS swaps (18M–50Y). The production curve — Swap_Pricer, Curve_Interface, and the CDS module all discount off this. |
| **Bootstrap_Lehman** | **Lehman-2002 methodology** on modern SOFR: short OIS (≤1Y) + **SR3 futures** (1Y–5Y, own quarterly-only chain) + OIS swaps (5Y–50Y). Compare col H against the OIS `Bootstrap` to see the futures-vs-OIS difference. |
| **Curve_Interface** | Shared `D(0,t)`: interpolates the SOFR curve to ANY date via piecewise-flat forward. Demo rows = standard quarterly CDS dates (20 Mar/Jun/Sep/Dec). The discounting engine for both Swap_Pricer and the CDS module. |
| **CDS_Parameters** | Credit assumptions — recovery (input, 40%), notional, standard coupon, direction. Recovery is an assumption, not bootstrapped. |
| **CDS_Quotes** | Single-name CDS par spreads: `BDP(ticker,"PX_LAST")` with a manual demo fallback so it runs immediately. One entity / seniority / clause. |
| **CDS_Schedule** | Quarterly premium & default schedule to 10Y. DF from the SOFR `Curve_Interface`, survival `Q(t)=Q_prev·exp(-λ·dt)` from the hazard curve, plus premium/accrual/protection PV factors per period. |
| **Hazard_Bootstrap** | Piecewise-constant hazard curve. Each λ defaults to the **forward-hazard approximation** (reprices within a few bp) and is **Goal-Seekable** (set the repricing-error cell to 0 by changing λ, top-down) for an exact fit. |
| **CDS_Pricer** | Prices any CDS off both curves: RPV01, protection & premium legs, par spread, PV, upfront, and the full first-order risk suite — **CS01, IR01/DV01, Rec01, jump-to-default**. |
| **CDS_Validation** | `Q(0)=1`, survival monotone, λ≥0, repricing error → 0, and the flat-curve check `s ≈ (1−R)·λ`. |

## The CDS module (credit curve on top of SOFR)

Modular by design — the SOFR curve stays the discounting engine; a **separate hazard-rate
bootstrap** sits on top, and the CDS pricer combines both. Nothing credit touches the Bootstrap
sheet.

- **Discounting** flows through `Curve_Interface` (the shared `D(0,t)`), so any change to
  fixings/futures/OIS propagates into every CDS PV automatically.
- **Survival** uses piecewise-constant hazards: `Q(0,t) = exp(−∫λ)`, the credit analogue of the
  SOFR curve's piecewise-flat forwards.
- **Par spread** `s = (1−R)·Σ D(T̄)·ΔQ / RPV01`, with `RPV01 = Σ α·D·Q + accrual-on-default`.
- **Hazard bootstrap:** each λ defaults to the forward-hazard approximation (validated to reprice
  within a few bp of the market spreads); for an exact fit run **Goal Seek** per pillar
  (Data → What-If → Goal Seek: set the repricing-error cell to 0 by changing λ, top row first).
  Fully self-contained in Excel — no VBA.

Recovery is an **input**: changing `R` re-fits every hazard even when spreads are unchanged
(`λ = λ(R)`).
| **Swap_Pricer** | Dynamic SWPM-style Fixed-vs-SOFR OIS valuation off the Bootstrap curve: NPV, par coupon, DV01/PV01, cashflow schedule. |
| **Bloomberg_S490_Validation** | Bloomberg's own USD SOFR curve (`YCSW0490 Index`) — benchmark, not an input. |
| **Conventions** | DES/FLDS convention grid for representative OIS (`USOSFRC/1/5/30`) + SWPM cashflow-export guidance. |
| **Fwd_Interp** | Zhou 2002 forward-curve interpolation: V1 piecewise-flat vs V2 piecewise-quadratic instantaneous forwards, both repricing the pillar DFs exactly. |

## The Fwd_Interp tab (Zhou 2002 forward interpolation)

Implements the paper's two forward-curve variants, both of which reprice the bootstrapped pillar
DFs **identically** (they share each segment's integral `∫f = −ln(DF_i/DF_(i-1))`):

- **Variant 1 — piecewise-flat instantaneous forward** (Lehman's default): the forward is constant
  over each segment, `f_i = ∫f / Δ`. A step curve — simplest, most stable, exact repricing.
- **Variant 2 — piecewise-quadratic instantaneous forward** `f(x)=a+bx+cx²`, continuous across
  pillars (adjacent segments share the "node forward") while still repricing exactly. The
  coefficients are **closed-form** from `(f_left, f_right, ∫f)` — no per-segment solver — because
  the bootstrap already fixes each segment's integral. Node forward = average of the two adjacent
  flat forwards. The `reprice check` column (∫quad − ∫f ≈ 0) confirms V2 preserves every DF.

The difference is intra-segment (compare **Flat fwd** vs **V2 fwd @ mid**), largest across the wide
long-end gaps (10Y→15Y→20Y→30Y) — exactly where V2's smoothing matters and V1's steps look coarse.

## How to use

1. On the **Instructions** tab, set the yellow parameter cells (valuation/snapshot date,
   `HIST_START`/`HIST_END`, price source — default `BGN`). Every date-driven pull references
   these named cells, so you set them once.
2. Let the sheets refresh. On **SOFR_Futures** and **Bloomberg_S490_Validation**, the column-A
   `BDS` anchors spill; trim the row count to the liquid strip using `VOLUME`/`OPEN_INT`.
3. Because your Bloomberg access may be temporary, **also export historical daily data** (the
   `BDH` block on SOFR_Fixings; snapshot the OIS/futures tabs) so the build is reproducible after
   the terminal is gone.
4. Feed the mid OIS quotes, futures-implied rates, and the fixing into the bootstrap
   (`openusdcurve` engine, `configs/sofr_market.yaml`); compare `P(0,T)`, zero and forward rates
   against the S490 tab.

## Ticker verification (checked July 2026)

| Ticker | Meaning | Verified against |
| --- | --- | --- |
| `SOFRRATE Index` | Published overnight SOFR fixing (NY Fed) | Bloomberg overnight cash indices |
| `USOSFRA`=1M, `B`=2M, `C`=3M … `USOSFR1`=1Y, `2`=2Y … `50`=50Y | USD fixed-vs-SOFR OIS strip | Clarus rates ticker database |
| `SFR` / `SER` roots → `SFRA` / `SERA` (active contract) | 3M (SR3) / 1M (SR1) SOFR futures | CME Group Bloomberg codes reference |
| `YCSW0490 Index` (curve **S490**) | Bloomberg's market USD SOFR swap curve | Bloomberg USD Bellwether Swap Indices |

Spot-check on the terminal (not publicly documentable): the weekly short-end codes
`USOSFR1Z/2Z/3Z` and the `USOSFR1F` (18M) code — these are the standard Bloomberg USOSFR
generic-tenor forms, but confirm they resolve under your entitlement.

## The Bootstrap tab

A self-contained, in-Excel single-curve OIS bootstrap from the `SOFR_OIS_Quotes` mid strip:

```
DF_n = (1 − S_n · A_(n-1)) / (1 + S_n · τ_n),   A_(n-1) = Σ τ_i · DF_i
```

where `S_n` is the par OIS mid, `τ` is ACT/360 between consecutive pillar maturities, and the
recursion runs top-to-bottom (each row references the rows above it). Outputs: **discount factor**
`P(0,T)`, **continuously-compounded zero rate** `z = −ln(DF)/T`, and the **simple forward** over
each pillar interval. Paste Bloomberg S490 zero rates into the yellow column J to get **Δz in bp**
— the same dealer-curve validation the Lehman paper performs.

**Three segments (Zhou 2002 structure), cutovers at 1Y and 5Y so nothing overlaps:**

- **Short (o/n–1Y)** — single-payment SOFR OIS: `DF = 1/(1 + S·τ0)`.
- **Middle (1Y–5Y)** — **SR3 futures**: the DF at 18M/2Y/3Y/4Y/5Y is interpolated off a futures
  chain (helper block, cols O–W) that discounts each contract's forward rate forward from the OIS
  1Y DF. Each contract discounts only its portion **beyond 1Y** (`τ eff`), so pre-1Y contracts
  pass through and the strip picks up exactly where the short OIS stops. A **convexity** column
  (T, in bp, default 0) lets you correct futures→forward.
- **Long (5Y–50Y)** — annual SOFR OIS swaps: `DF = (1 − S·A_prior)/(1 + S·τ_coupon)`, with the
  annuity (col I) built from the short + futures-derived DFs.

Validated numerically: the futures-built mid DFs are monotonic and land within a few bp of the
pure-OIS values (they price the same market). Payment convention is respected throughout — OIS
≤1Y pay one coupon, ≥18M pay annually, and the annuity books `τ·DF` only at true annual coupon
dates. For the fully exact interpolated-annuity build, use the `openusdcurve` Python engine.

## The Swap_Pricer tab

A dynamic SWPM replica: value any Fixed-vs-SOFR OIS off the bootstrapped curve. Inputs (yellow):
direction, notional, effective date, tenor, coupon (defaults to par → NPV ≈ 0). It interpolates
discount factors **log-linearly** off the Bootstrap curve, builds the annual fixed-leg cashflow
schedule, and reports **NPV, par coupon, DV01/PV01**, and both leg NPVs — SWPM-style. The float
leg uses the single-curve telescoping identity `PV = N·(DF_eff − DF_mat)`.

Cross-checked against the SWPM screen for a 10mm 1Y swap: **par coupon ≈ 4.02%, net NPV = 0,
PV01 ≈ 974** (SWPM 973.78). SWPM's leg NPVs read ~9.995mm rather than 10mm because it discounts
from the curve date to a valuation date 2 business days after effective; here effective = spot so
`DF(effective) = 1` exactly. Net NPV, par coupon and DV01 are unaffected.

## Field-mnemonic notes (verified live on terminal)

- **Futures:** `FUT_CONTRACT_DATE` is **not** a valid Bloomberg field — replaced with
  `FUT_DLV_DT_FIRST` (the SR3/SR1 reference-quarter start, i.e. the future's accrual start).
  `SECURITY_DES`, `PX_LAST/BID/ASK`, `FUT_MONTH_YR`, `LAST_TRADEABLE_DT`, `VOLUME`, `OPEN_INT`
  all resolve; implied rate = `100 − PX_LAST`.
- **Conventions:** `USOSFR… BGN Curncy` are *rate-quote* securities and do not carry leg-level
  static fields, so only `CRNCY`, `CALENDAR_CODE`, and `DAY_CNT_DES` are pulled via `BDP`. All
  other conventions (floating day count, pay frequency, payment lag, EOM, compounding, lookback)
  are captured from `DES`/`SWPM` on an actual swap — the tab lists the USD-SOFR-OIS
  market-standard defaults as the expected values. Note: Bloomberg returns `#N/A Invalid Field`
  as **text**, so `IFERROR` cannot suppress it — using valid mnemonics (not wrapping) is the fix.

## Notes / caveats

- **Price source consistency:** the `PX_SRC` parameter (default `BGN`) is applied across every
  OIS ticker — don't mix `BGN` and `CBBT` unless deliberate.
- **Convention field mnemonics are best-guess and entitlement-dependent.** Where a Conventions
  cell shows `N/A — use FLDS`, run `FLDS <GO>` on that security to find the enabled field.
- **Licensing:** Bloomberg data is licensed — do not redistribute pulled values or commit a
  populated copy. Keep saved snapshots out of any public repo (this repo's `.gitignore` already
  excludes pulled data).
