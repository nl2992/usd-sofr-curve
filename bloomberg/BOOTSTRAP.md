# USD SOFR Curve — Bootstrap Methodology

How the `Bootstrap` tab turns market quotes into a discount curve `P(0,T)`, and how the rest of
the workbook (Curve_Interface, Swap_Pricer, the CDS module) consumes it.

## 1. Three-segment structure (Zhou 2002)

The curve is built in three maturity segments with **cutovers at 1Y and 5Y so nothing overlaps** —
each maturity is covered by exactly one instrument type:

| Segment | Range | Instruments | Discount factor |
| --- | --- | --- | --- |
| **Short** | o/n → 1Y | SOFR fixing + short OIS (1W…1Y) | single payment: `DF = 1 / (1 + S·τ0)` |
| **Middle** | 1Y → 5Y | 3M SOFR (SR3) futures | forward-chained from the 1Y DF (see §3) |
| **Long** | 5Y → 50Y | annual SOFR OIS swaps | `DF = (1 − S·A_prior) / (1 + S·τ_cpn)` |

This is the modern (SOFR) analogue of the Lehman/Zhou single-curve build: deposits → futures →
swaps becomes short-OIS → SR3 futures → OIS swaps.

## 2. Payment conventions (why the formulas differ by segment)

- **Sub-1Y OIS pay a single coupon** at maturity, so their DF comes straight from the money-market
  identity `DF = 1/(1 + S·τ0)`, with `τ0 = ACT/360(spot, maturity)`.
- **OIS ≥ 18M pay annual coupons.** The par condition `1 − DF(T_n) = S·Σ τ_i·DF(T_i)` is inverted
  sequentially: `DF_n = (1 − S·A_(n-1)) / (1 + S·τ_cpn)`, where `A_(n-1) = Σ τ_i·DF_i` accumulates
  **only at annual coupon dates** (column *Ann.add* books `τ·DF` on annual pillars and 0 elsewhere,
  so 2Y's annuity correctly sees 1Y — not 18M or the monthly points).

Every input reprices to par to ~0 bp (verified numerically 1W–50Y).

## 3. The middle segment (SR3 futures)

Each SR3 contract implies a forward rate `f = (100 − price)/100 − convexity`, where the
**convexity adjustment** `CA = ½·σ²·t1·t2` (σ = short-rate normal vol, an input, default 100 bp/yr)
corrects futures → forwards. The chain discounts each contract's forward forward from the OIS 1Y
discount factor, counting only the portion **beyond 1Y** so pre-1Y contracts pass through and the
strip picks up exactly where the short OIS stops. The DF at 18M/2Y/3Y/4Y/5Y is then interpolated
off that futures chain.

**Robustness guard.** Futures and OIS price the same market, so a good futures DF sits within a few
hundredths of the OIS DF. The middle DF is therefore
`IF(|DF_futures − DF_OIS| < 0.03, DF_futures, DF_OIS)`: when the SR3 chain is healthy the futures
drive the middle (as intended); if the chain returns bad/mixed contracts the DF is rejected and the
curve falls back to the **robust OIS bootstrap**, so the whole curve past 1Y can never blow up. The
*DF source* column reports which is live per pillar.

## 4. Interpolation (piecewise-flat instantaneous forward)

Between pillars the instantaneous forward is held constant, i.e. `ln P` is linear in time. For
`t ∈ (T_(j-1), T_j]`:

```
f_j    = −[ln P(T_j) − ln P(T_(j-1))] / (T_j − T_(j-1))
P(0,t) = P(0, T_(j-1)) · exp(−f_j · (t − T_(j-1)))
```

This is the SOFR curve's baseline method and the shared discounting used everywhere. The
`Fwd_Interp` tab also implements the paper's smoother **piecewise-quadratic** variant, which
reprices the same pillar DFs but makes the forward continuous across pillars.

## 5. Outputs and consumers

- **Discount factor** `P(0,T)` (col H) — the deliverable curve. **Zero rate** `z = −ln(DF)/T`
  (ACT/365). **Forward** over each pillar interval.
- **Curve_Interface** exposes `P(0,t)` at *any* date via the §4 formula (demo: standard quarterly
  CDS dates). It is the single discounting engine for both the **Swap_Pricer** and the **CDS
  module** (schedule DFs, hazard bootstrap, CDS pricing), so any change to fixings / futures / OIS /
  interpolation flows through automatically.
- **Validation:** paste Bloomberg's S490 zero rates into `Bloomberg_S490_Validation` (col D) to get
  `Δz` per tenor — single-digit bp confirms agreement with the street curve.
