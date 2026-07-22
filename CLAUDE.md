# CDS module — deliverables to reproduce the CDSW screen

Source of truth: **Bloomberg CDS Model white paper (B-Model), pp.4–11**, photographed in
`CDS/IMG_7801–7808.HEIC`. Equation numbers below are that paper's.
Secondary: Help Desk ticket **H#1330731572** (the BDP pull).

**O'Kane & Turnbull (2003), "Valuation of Credit Default Swaps", Lehman Brothers QCR** —
`OKaneTurnbull2003Lehman.pdf` in this repo, read. Equations cited as O-T (n).

The two papers agree on everything structural. O-T §9 names the method "bootstrapping", assumes
piecewise-flat hazard, solves each maturity in turn by **one-dimensional root search — footnote 8
says bisection or Newton-Raphson** — and holds the hazard flat beyond the last maturity. That is
`Hazard_Solver` line for line.

Where they appear to differ, they do not:
- O-T (9) breakeven `S = Protection / RPV01` has **no** accrued-interest netting; B-Model (3.3)
  does. O-T (9) is written for a **new** contract at `t_V = t_0`, where nothing has accrued.
  (3.3) is the general seasoned case with a settlement lag. D2 follows (3.3), which is correct
  for CDSW.
- O-T `Z(t_V,t_n)` is "the Libor discount factor... bootstrapped in the currency of the default
  swap". For USD today that is the SOFR/RFR curve, which is what B-Model p.6 states outright.

Deliverable is the single workbook
`~/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg.xlsx`.

---

## Non-negotiable: the DF component

p.6 — *"The B-Model in CDSW also uses RFR curves by default for these currencies"* (USD, EUR,
JPY, GBP, CHF, AUD).

So `D(t)` **is** our bootstrapped SOFR curve. Every credit quantity routes through
`Curve_Interface`. Nothing in the CDS module may use a flat rate, its own curve, or a
hardcoded DF. The curve is already validated to S490 at 0.397bp — that validation is what
gives the CDS numbers standing.

Also p.6: the pricing date `T` is the **origin** of both `p_s(t)` and `D(t)`.

Chain, verified numerically (re-check after any change that touches discounting):

```
SOFR_OIS_Quotes!H   BDP mid, falls back to the frozen S490 mids
  -> Bootstrap!H     = $D$4/(1+(E/100)*D)   short   |  (1-S*A)/(1+S*tau)  long
  -> Curve_Interface!L9 = Bootstrap!H8      the shared D(0,t) grid
  -> CDS_Schedule!G,H    pay-date and mid-period DFs
  -> CDS_Parameters!B22  D(T_s) at settlement
```

Implied zeros at CDS pay dates track the validated curve: 3.76% @2M, 4.06% @1Y, 4.05% @5Y,
4.19% @10Y.

---

## Already correct — do not "fix"

Verified against the paper, leave alone:

| what | where | paper |
|---|---|---|
| Piecewise-constant hazard, `p_s(t) = p_s(t_{i−1})·exp(−h_i·(t−t_{i−1}))` | `Hazard_Solver` | p.5 |
| Sequential 1-D root-find per tenor, retaining earlier `h₁…hᵢ` | `Hazard_Solver` | §4 p.8 |
| `Market Value = Protection Leg − Premium Leg` | `CDS_Pricer!B16` | (3.1) |
| `Coupons-in-Survival = A·C·Σ Δ(T_{i−1},T_i)·p_s(T_i)·D(T_i)` | `CDS_Schedule!M` | (3.4) |
| Accrued days include the pricing date (the `+1`) | `CDS_Pricer!B34` | p.7 |
| `h ≥ 0` enforced | bisection bracket `[0,3]` | p.8 |
| ACT/360 premium, 100/500bp coupon, IMM 20th | `CDS_Parameters` | p.6 |
| `D(t)` = RFR/SOFR curve | `Curve_Interface` | p.6 |

Our sequential bisection *is* the documented B-model stripper. That was confirmed, not assumed.

---

## Deliverables

### D1 — Settlement date `T_s`   ✅ DONE
p.6: settlement is **T + 3 business days**. Not currently modelled anywhere.

- Add `T_s` to `CDS_Parameters`, business-day rolled.
- Expose `D(T_s)` off `Curve_Interface`.

### D2 — Par spread per (3.3)   ✅ DONE
```
S = Protection Leg / (Premium Leg − Accrued Interest·D(T_s))|_{C=1bp}
```
We currently use `Protection / RPV01` with **no accrued-interest netting**. The denominator is
PV01 = premium leg *net of accrued interest* at C=1bp, discounted from settlement.

- Fix `CDS_Pricer!B15` and the same expression inside `Hazard_Solver`'s objective.
- Re-verify the stripper still reprices its inputs after the change.

### D3 — Upfront per (3.2)   ⚠️ PARTIAL
```
Upfront = Market Value / D(T_s) + Accrued Interest
```
We use `(S−C)·RPV01`, missing both the compounding to settlement and the accrued rebate.

Delivered: `CDS_Pricer!B20` Market Value (3.1), `B26` Accrued Interest, `B27` Upfront (3.2).
Upfront reads 9.1e-12 at `C = S`, which is the sign check on the accrued term.

**Not delivered:** how (3.2)'s single Upfront splits into CDSW's Principal / Accrued / Cash.
The paper gives the total only. The split was attempted and would not close against the
reference screen without assuming a convention the paper does not state, so the old split
stands with a red flag on `CDS_Pricer!A38`. **This is the acceptance test for the module** —
one live CDSW comparison at close fixes it.
Target: CINDBK 5Y 07/21/26 → pts upfront −1.97596265, price 101.97596265, principal −197,597,
accrued −8,333 (30d), cash −205,930, def exp 6,197,596.

### D4 — Risk by re-stripping, not analytic   ⬜ NEXT, largest build
p.9: *"Spread DV01 and IR DV01 represent the change in the value of the transaction as a result
of a parallel shift of 1bp in the CDS curve or interest rate curve respectively"*, obtained by
**rerunning survival probability stripping**.

Ours are first-order analytic (`CS01 = N·RPV01·1bp`) and `Rec01` is explicitly sticky-hazard.
That is not what CDSW reports.

- Add three bumped stripper runs: CDS curve +1bp, SOFR curve +1bp, recovery +1%.
- Each re-solves all six hazards, then reprices. Same pattern as `Curve_Solver`.
- Plus jump-to-default (sensitivity to immediate default), already present.
- These are hidden working sheets.

### D5 — CLOSED, no action. Keep Z(mid).
Measured against the right reference and the answer reversed, so recording it rather than
leaving a trap.

Bloomberg (3.5)/(3.6) integrate `D(t)` at the **default time**, analytically on fine segments.
`Z(mid)` is the second-order approximation to that; O-T's `Z(t_n)` is first-order. Stripping a
5Y at 95bp, r=4%, R=40%, against the exact continuous integrals:

```
EXACT (3.5)/(3.6)                lambda = 0.01575435
ours: quarterly, Z(mid)          lambda = 0.01575454    +0.12 bp
O-T (5)+(7) M=12, Z(end)         lambda = 0.01578056   +16.64 bp
O-T M=4,        Z(end)           lambda = 0.01583335   +50.15 bp
Z(mid), M=12 protection          lambda = 0.01575444    +0.05 bp
```

**Do not adopt O-T (5)'s `Z(t_n)`** — it is 140x worse against Bloomberg. Do not refine the grid
either; M=12 buys 0.07bp of hazard.

The earlier version of this item recommended switching to `Z(t_n)` because it benchmarked against
O-T M=12 as if that were truth. It is not: it is another discretisation, and a coarser one. When
O-T and the B-Model disagree on a *numerical convention*, follow the B-Model — CDSW is the thing
being replicated. O-T remains the better reference for derivations and for why the model is built
the way it is.

### D6 — IMM roll: last accrual date is NOT rolled   ⬜
p.6: `T_i` are IMM dates *"rolled to the next business day **except the last date which is not
rolled**. Payment dates are the same as accrual end dates with an exception at maturity... The
payment date at maturity rolls to the next business day, but the last accrual end date `T_M`
does not."*

So maturity needs **two** dates: accrual end (unrolled) and payment (rolled). We roll everything.

- Split them in `CDS_Schedule`.

### D7 — Protection period   ⬜
p.7: protection starts immediately; full day count for protection and coupon is `(T_M − T + 1)`.
Protection seller receives the **full** coupon for the previous period on the first coupon date.

- Check the front stub in `CDS_Schedule` row 7 against both rules.

### D8 — Strippability check   ⬜
p.9: calibration fails if the target spread is unreachable; common with inversions. Lower bound
from the credit triangle:
```
S_{i+1} ≥ S_i · T_{M_i} / T_{M_{i+1}}
```
- Add to `CDS_Validation` as a pass/fail per tenor, so a failed strip is visible rather than a
  bisection silently pinning at a bracket end.

### D9 — Default probabilities on the screen   ⬜
p.6: `p_d(t) = 1 − p_s(t)` is what CDSW reports as "probabilities".

- Add the `Prob` column to the CDSW panel; the reference screen shows 0.0446 at 5Y.

### D10 — Forward valuation (curve date ≠ valuation date)   ⬜ lowest
p.10: *"The valuation date in CDSW can be set later than the curve date... the model produces a
forward value for the deal, using survival probabilities conditional on survival until the
valuation date."*

The reference CDSW screen has Curve Date 07/21/26, Valuation 07/21/26; the SWPM screen showed
07/21 vs 07/23. Same mechanism explains the leg-NPV basis difference noted on `Swap_Pricer!A32`.

- Add a valuation date separate from the curve date, with survival conditioned on it.
- Lower priority: only matters when the two differ.

---

## Order

D1 → D2 → D3 done; D5 closed with no action.

Remaining: **D6/D7** (roll conventions — cheap, touch every date, do first) → **D4** (risk by
re-stripping — largest build) → **D8/D9** (small, visible) → **D10** (only bites when curve date
≠ valuation date).

Blocked on market data, not on code: D3's split and every CDS number, until the ticker pulls
resolve at close.

## Rules

- **Verify before writing.** Model it in Python against the CDSW targets first; only patch the
  workbook once the numbers land. Every change so far was checked this way and it caught real
  bugs (the merged cell over `Swap_Pricer!K39:L40`, the `F22` spot/settle mismatch).
- **After every patch:** recalculate and confirm `Bootstrap!U76 = 32/32`, `U77 ≈ 0.397bp`,
  `Hazard_Bootstrap!J ≈ 0`, and the error count has not risen.
- **No fitted parameters.** Inputs are pulled quotes or documented conventions. Do not tune
  anything to make a target match — that happened once with the o/n stub and had to be undone.
- **One workbook.** No second file.

## Verification

Pull list for the terminal: **`TEST_CASES.md`**. Minimum viable set is T1 (S490 matched
pair), T3 (real CDS spreads) and T4 (CDSW full screen). Work is on hold until those land.

## Open

- **`CDS_Quotes!E` still holds demo spreads** (40/55/70/95/115/135), not market. Every CDS number
  is provisional until the `CDS_SPREAD_TICKER_nY` pulls resolve on a terminal.
- **No external validation exists for the hazard curve.** Help Desk: Bloomberg exports neither
  zero-coupon CDS spreads nor stripped hazard rates. The only check is CDSW output for a known
  trade — which makes D3's target list the acceptance test for the whole module.
- p.11: spread-vs-upfront calibration targets cannot both be honoured simultaneously (the ISDA
  converter assumes globally constant hazard). **We calibrate to spreads.** Note it, don't fix it.

---

## Benchmark available: O-T Figure 10

A fitted curve, R=40%, that the stripper should be able to reproduce given the same discount
curve:

| Term | Hazard | Market bp | Model bp | Protection Leg | Risky PV01 |
|---|---|---|---|---|---|
| 6M | 1.6832% | 100 | 100 | 0.4941% | 0.4941 |
| 1Y | 2.0203% | 110 | 110 | 1.0825% | 0.9841 |
| 2Y | 2.1950% | 120 | 120 | 2.3061% | 1.9218 |
| 3Y | 3.0838% | 140 | 140 | 3.9264% | 2.8046 |
| 5Y | 2.8126% | 150 | 150 | 6.6329% | 4.4219 |
| 7Y | 3.2054% | 160 | 160 | 9.3720% | 5.8575 |
| 10Y | 3.0386% | 165 | 165 | 12.7162% | 7.7068 |

The table is internally consistent — `Protection / RPV01` returns the model spread to 4dp at
every tenor, which is O-T (9). **The paper does not publish the discount curve used**, so the
hazards themselves cannot be reproduced exactly; treat it as a shape and magnitude check, not a
unit test. The relationship `Protection / RPV01 = quoted spread` *is* testable and already holds
in our stripper to ~1e-6 bp.
