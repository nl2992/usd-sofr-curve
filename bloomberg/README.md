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
| **SOFR_OIS_Quotes** | Full `USOSFR…` OIS strip 1W–50Y: bid/ask/last + mid `=(bid+ask)/2` (falls back to last); recommended bootstrap set flagged. |
| **Bootstrap** | Live single-curve OIS bootstrap from the OIS mids → discount factors, zero & forward rates, plus a Δz-vs-S490 column. |
| **Swap_Pricer** | Dynamic SWPM-style Fixed-vs-SOFR OIS valuation off the Bootstrap curve: NPV, par coupon, DV01/PV01, cashflow schedule. |
| **Bloomberg_S490_Validation** | Bloomberg's own USD SOFR curve (`YCSW0490 Index`) — benchmark, not an input. |
| **Conventions** | DES/FLDS convention grid for representative OIS (`USOSFRC/1/5/30`) + SWPM cashflow-export guidance. |

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

**Payment convention (fixed):** OIS ≤ 1Y pay a **single** coupon at maturity, so those points use
`DF = 1/(1 + S·τ0)`; OIS ≥ 18M pay **annual** coupons, so `DF = (1 − S·A_prior)/(1 + S·τ_coupon)`
where `A_prior` accumulates `τ·DF` **only at annual coupon dates** (column I books zero on the
sub-annual and 18M rows, so 2Y's annuity correctly sees 1Y — not 18M or the monthly points). This
was the fix for an earlier version that treated every pillar as a coupon date and mis-priced the
front ~7bp. **Every input now reprices to par to ~0 bp** (validated numerically 1W–50Y); the
sparse long end (12–50Y) still uses the pillar grid as its coupon schedule. For the fully exact
interpolated-annuity build, use the `openusdcurve` Python engine (`configs/sofr_market.yaml`).

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
