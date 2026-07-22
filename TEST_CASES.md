# Bloomberg pull list — what to capture at close

Verification set for `USD_SOFR_Curve_Bloomberg_Pricer.xlsx`. Ordered by what blocks what.

## The one rule

**Capture each test case in a single sitting and note the timestamp.** Every mismatch this
build has produced traced to comparing a curve built from one snapshot against a target from
another. T1 and T3 in particular are matched pairs — inputs and outputs from the *same instant*
or they prove nothing.

---

## P1 — Blocks everything downstream

### T1. S490 matched pair  ⟵ *most important*
`YCSW0490 <GO>` → **Curve Construction** tab, then **Curve Analysis** tab, without refreshing between.

| capture | fields |
|---|---|
| Construction | Swap Rates panel: Term, Bid, Ask for all 32 tenors (1WK→50YR) |
| Construction | Interpolation, Settle Date, Curve Side, Shift, PCS, Curve Date |
| Analysis | Date, Market Rate, Zero Rate, Discount — all 32 rows |

**Verifies:** the whole bootstrap. **Accept:** `Bootstrap!U76 = 32/32`, `U77 ≤ 0.5bp`, `U78 ≤ 5e-05`.
Currently 0.397bp / 1.88e-05 on the 07/21 capture.
**Watch:** mid = (bid+ask)/2 must equal the Market Rate column. If it doesn't, the curve side or
pricing source differs from what we assume.

### T2. BDP pre-flight
On the terminal with the workbook open, `Bootstrap!G4 = Live (BDP)`.

| check | where | fail condition |
|---|---|---|
| Units | `SOFR_OIS_Quotes!H` vs `J` | H ≈ 100× J → BDP returning decimals, everything breaks |
| Tickers | `SOFR_OIS_Quotes!E:G` | any `ERROR PLZ FIX` → bad ticker for that tenor |
| Live vs fixed | `SOFR_OIS_Quotes!U41` | large + non-zero mean at `C82` → wrong pricing source |

**Verifies:** the dynamic path. Only meaningful if column `T` is populated — two modes agreeing
with `T` blank proves nothing.

### T3. Real CDS par spreads  ⟵ *everything credit is fake without this*
For one liquid name, e.g. the CINDBK reference already on file.

- `CDS_SPREAD_TICKER_1Y/3Y/5Y/7Y/10Y` results, and the quoted spread per tenor
- Recovery rate the screen assumes, seniority, restructuring clause, currency
- Whether a 2Y ticker exists (our 2Y row is speculative)

**Verifies:** replaces the demo spreads (40/55/70/95/115/135), which are placeholders.
**Accept:** `Hazard_Bootstrap!J` stays ~1e-6 bp and hazards stay positive and monotone-ish.

---

## P2 — Verifies what is already built

### T4. CDSW full screen  ⟵ *the acceptance test for the CDS module*
`CDSW <GO>` on the T3 entity, 5Y.

Capture **inputs**: notional, currency, trade date, valuation date, curve date, cash-settle date,
1st accrual start, 1st coupon, maturity, coupon, recovery, day count, frequency, buy/sell,
traded spread, curve used.

Capture **every output**: Pts Upfront, Price, Principal, Accrued (+ the day count shown),
Cash Amount, Spread DV01, IR DV01, Rec Risk (1%), Def Exposure, and the **Prob** column.

**Verifies:** D3, and it is the *only* external check the credit side can ever have — Bloomberg
exports no hazard curve (Help Desk H#1330731572).
**Resolves:** how (3.2)'s single Upfront splits into Principal / Accrued / Cash. That mapping is
currently unverified and flagged red on `CDS_Pricer!A38`.
**Accept:** Upfront, Price, Cash Amount within a basis point of ours.

### T5. SWPM across the curve
`SWPM <GO>`, fixed-vs-SOFR, 10mm, same curve date.

Tenors **1W, 1Y, 5Y, 10Y, 20Y, 30Y, 50Y**. Per deal: Par Cpn, NPV, PV01/DV01, Accrued,
Leg 1/Leg 2 NPV, plus Curve Date and Valuation.

**Verifies:** the swap pricer. 20Y+ specifically — that is where the merged-cell hole in the
curve grid hid, and 1W exercises the sub-annual path added for the maturity override.
**Accept:** par reprices its own input ≤ 0.2bp at every tenor.
**Note:** leg NPVs will differ from ours by construction — SWPM discounts curve date → valuation
date. Capture them anyway; they are what would let us close D10.

---

## P3 — Resolves the open deliverables

### T6. Maturity on a weekend  → D6
A CDS whose maturity IMM date falls on a Saturday or Sunday.

Capture the **accrual end date and the payment date separately**. White paper p.6 says the
payment date rolls but the last accrual end date `T_M` does **not**. We currently roll both.

### T7. Curve date ≠ valuation date  → D10
One CDSW with Valuation set later than Curve Date, everything else as T4.

Capture the same output set. Also explains the SWPM leg-NPV basis difference noted on
`Swap_Pricer!A32`.

### T8. An inverted or distressed curve  → D8
Any name where a shorter spread exceeds a longer one (`S_{i+1} < S_i`).

Capture the term structure and whether CDSW strips it or errors. White paper p.9 gives the
strippability bound `S_{i+1} ≥ S_i · T_i / T_{i+1}`.
**Verifies:** our bisection fails *visibly* rather than pinning silently at a bracket end.

### T9. Risk by re-stripping  → D4
On the T4 deal, re-read Spread DV01, IR DV01 and Rec Risk after a **+1bp shift applied on the
screen** (the Shift field), rather than reading the reported sensitivities.

**Verifies:** p.9 says CDSW obtains these by rerunning the stripper. Ours are analytic
first-order. This gives the target to build D4 against.

### T10. Front stub  → D7
A CDS traded **between** IMM dates, so the first period is a stub.

Capture accrued days, the first coupon amount, and protection start. p.7: protection starts
immediately, day count is `(T_M − T + 1)`, and the seller gets the **full** coupon for the
previous period on the first coupon date.

---

## Quick reference — what each screen gives

| screen | gives | test |
|---|---|---|
| `YCSW0490` Construction | swap rates bid/ask, curve settings | T1 |
| `YCSW0490` Analysis | zero + discount by date | T1 |
| `CDSW` | CDS pricing, settlement, risk, probabilities | T4, T6, T7, T9, T10 |
| `SWPM` | swap par, NPV, DV01 | T5 |
| `CDSD` / curve screen | CDS term structure per entity | T3, T8 |
| Excel `BDP`/`BDS` | live feed into the workbook | T2 |

## Minimum viable set

If time is short: **T1 + T3 + T4**. That re-validates the curve on fresh data, replaces the fake
credit inputs, and settles the one open question in the CDS pricer. Everything else can wait.
