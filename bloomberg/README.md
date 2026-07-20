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

## Notes / caveats

- **Price source consistency:** the `PX_SRC` parameter (default `BGN`) is applied across every
  OIS ticker — don't mix `BGN` and `CBBT` unless deliberate.
- **Convention field mnemonics are best-guess and entitlement-dependent.** Where a Conventions
  cell shows `N/A — use FLDS`, run `FLDS <GO>` on that security to find the enabled field.
- **Licensing:** Bloomberg data is licensed — do not redistribute pulled values or commit a
  populated copy. Keep saved snapshots out of any public repo (this repo's `.gitignore` already
  excludes pulled data).
