# KRW IRS bootstrap

Quarterly-fixing KRW IRS curve, 3m to 30Y, built to Kevin's Date/Year/Swap/Zero/
DF/Forward spec. Three sheets, in order:

1. **Quotes** — the 15 quoted tenors and their swap rates. 3m (2.91) is the
   fixing and is treated as the zero rate. `tau` is a single input cell.
2. **Interpolation** — every quarter from 3m to 30Y, with its swap rate linearly
   interpolated in time between the two bracketing quotes. A quoted tenor keeps
   its own rate. This sheet is where the gaps are filled: 20Y to 30Y is 40 rows,
   handled exactly like the one row between 1Y and 15m.
3. **Bootstrap** — DF, zero and forward off the quarterly grid.

## The whole method in two lines

    3m (money market):   DF = 1 / (1 + z*tau)
    every swap:          DF(T) = (1 - S*tau * sum of earlier DFs) / (1 + S*tau)

A par swap is worth zero, so fixed leg = float leg, and the float leg telescopes
to 1 - DF(T). That leaves one unknown per row - the new DF - using every DF
already found. Zero = (1/DF - 1)/t. Forward = (DF_prev/DF - 1)/tau.

## Why the big gaps are not special

The par equation for any node needs a DF at every earlier coupon date. Fill each
quarter's swap rate first (the Interpolation sheet) and every one of those dates
becomes a node with a single unknown. A 40-quarter gap is then 40 sequential
one-line solves, no different from one. Leaping straight from 20Y to 30Y would
leave 39 unknowns in one equation; walking every quarter leaves one.

## Simplifications, to revisit with Kevin

- `tau = 0.25` flat, not ACT/365 actual day counts.
- Linear interpolation on the swap rate. Interpolating on the zero rate, or
  log-linear on DF (piecewise-flat forward), would give slightly different
  intermediate points and smoother forwards.

Live: change any quote on the Quotes sheet and the whole curve reprices.
Verified against a hand bootstrap - every DF agrees to 7 decimals.

## Building it by hand (the Book1 rebuild)

If you are rebuilding the three sheets by hand rather than copying them, the live
formulas are below. Layout used here — note it differs from KRW_IRS_Bootstrap.xlsx,
which has year and rate the other way round:

**Quotes** — `A` tenor label, `B` Coupon (rate %), `C` Year. Data rows 5-19.
`tau` in `B2` (0.25).

**Interpolation** — grid starts row 3, `A` = Year (0.25, 0.5, … to 30).
Helper `G` finds the bracketing quote row; `E` is the interpolated swap rate.

```excel
G3  =MATCH($A3,Quotes!$C$5:$C$19,1)
D3  =IF($C3=$B3,0,($A3-$B3)/($C3-$B3))     ' weight, if you keep lo/hi columns
E3  =INDEX(Quotes!$B$5:$B$19,$G3)+$D3*(INDEX(Quotes!$B$5:$B$19,MIN($G3+1,15))-INDEX(Quotes!$B$5:$B$19,$G3))
```

`E` reads: lo rate + weight × (hi rate − lo rate). Rates come from `Quotes!$B`
because in this layout `B` is the coupon and `C` is the year. `MIN($G3+1,15)`
caps the "hi" at the last quote so 30Y does not run off the end.

**Bootstrap** — starts row 3, `A` Year, `B` Swap rate, `C` DF, `D` Zero, `E` Fwd.

```excel
A3  =Interpolation!A3
B3  =Interpolation!E3
```

Row 3 is the 3m money-market row — type these once:

```excel
C3  =1/(1+(B3/100)*Quotes!$B$2)
D3  =(1/C3-1)/A3*100
E3  =(1/C3-1)/Quotes!$B$2*100
```

Row 4 down are all identical — type in row 4 and fill to 122:

```excel
C4  =(1-(B4/100)*Quotes!$B$2*SUM($C$3:C3))/(1+(B4/100)*Quotes!$B$2)
D4  =(1/C4-1)/A4*100
E4  =(C3/C4-1)/Quotes!$B$2*100
```

`SUM($C$3:C3)` — the anchor stays, the end slides, so each DF sums every earlier
DF; that is the bootstrap chaining. `Quotes!$B$2` is tau, so changing it reprices
the whole curve.

All formulas verified against the hand bootstrap: every DF to 7 decimals, zero
errors.

### Watch out: is the swap rate a percent or a fraction?

This is the single trap when rebuilding. It decides whether the formulas carry a
`/100` or not, and getting it wrong makes every DF sit just under 1 (a 10Y DF of
0.996 instead of 0.657) — plausible-looking and 100x wrong.

- If the Swap rate cell holds **2.91** (plain number, format `0.0000`): divide by
  100 in the formulas — `(B/100)`.
- If it holds **0.0291** (i.e. the cell is *percent-formatted* and shows `2.910%`):
  the value is already a fraction, so **drop the `/100`**.

`KRW_IRS_Bootstrap.xlsx` stores 2.91 (plain), so its formulas use `/100`. A
percent-formatted rebuild stores 0.0291 and must not. Tell them apart by clicking
the cell: the formula bar shows the underlying value.

Bootstrap formulas for the **percent-formatted** case (B already a fraction).
Row 3 is the 3m money-market row; row 4 down is the drag-down:

```excel
' row 3
A3  =Interpolation!A3
B3  =Interpolation!E3
C3  =1/(1+B3*Quotes!$B$2)
D3  =(1/C3-1)/A3
E3  =(1/C3-1)/Quotes!$B$2

' row 4, fill down to row 122 (30Y)
A4  =Interpolation!A4
B4  =Interpolation!E4
C4  =(1-B4*Quotes!$B$2*SUM($C$3:C3))/(1+B4*Quotes!$B$2)
D4  =(1/C4-1)/A4
E4  =(C3/C4-1)/Quotes!$B$2
```

Zero (D) and Future (E) now hold fractions — format them as `%`, or append `*100`
to those two lines only (never to the DF line). Sanity check after filling down:
DF(10Y) = 0.657, zero climbs to 6.72% at 30Y, forward turns down to 2.64%. If
DF(10Y) is still near 0.996 the `/100` is still in the DF formula.

Verified to 30Y against the hand bootstrap: every DF to 7 decimals, zero errors.

---

## Hedging — Bloomberg tickers and the "is it worth it" test

KRW IRS vs 91-day CD, generic composite tickers. **Confirm each with `<ticker> Curncy DES <GO>`
and eyeball the market on `ALLQ <GO>` before wiring** — generic tickers occasionally re-map.

| Tenor | Ticker | Role in the hedge |
|-------|--------|-------------------|
| 3M CD | `KWCDC Curncy`   | fixing / front — **dust, do not hedge** |
| 1Y    | `KWSW1 Curncy`   | folds into 2Y |
| 2Y    | `KWSW2 Curncy`   | **hedge** — receive fixed |
| 3Y    | `KWSW3 Curncy`   | **hedge** — receive fixed |
| 4Y    | `KWSW4 Curncy`   | folds into 5Y |
| 5Y    | `KWSW5 Curncy`   | **hedge** — pay fixed |
| 7Y    | `KWSW7 Curncy`   | optional split (usually skip) |
| 10Y   | `KWSW10 Curncy`  | **hedge** — pay fixed |
| 15Y   | `KWSW15 Curncy`  | — (past maturity) |
| 20Y   | `KWSW20 Curncy`  | — |

Reference curve: `YCSW0074 Index` (KRW IRS). CD fixing history: `KWCDC Curncy`.

### BDP pulls (Excel add-in) — bid, ask, spread in bp

```excel
=BDP("KWSW7 Curncy","PX_BID")
=BDP("KWSW7 Curncy","PX_ASK")
=(BDP("KWSW7 Curncy","PX_ASK")-BDP("KWSW7 Curncy","PX_BID"))*100      ' spread, bp
```

Average spread over the last month (proves it isn't a one-off wide quote):

```excel
=AVERAGE(BDH("KWSW7 Curncy","PX_ASK",TODAY()-30,TODAY()))
   -AVERAGE(BDH("KWSW7 Curncy","PX_BID",TODAY()-30,TODAY()))
```

### Cost to execute a hedge (cross half the bid/offer, valued through the annuity)

    Cost_KRW = Notional × Annuity(T) × (spread_bp / 2) × 0.0001

```excel
' 7Y example: ~22bn notional, annuity 6.073 from the Bootstrap sheet
=22e9 * 6.073 * (BDP("KWSW7 Curncy","PX_ASK")-BDP("KWSW7 Curncy","PX_BID"))/2/100
```

### Decision rule

> **Hedge a bucket only if:  risk removed (KRW/bp)  >  ½ × bid-offer(bp) × notional × annuity.**

- **2Y / 3Y / 5Y / 10Y** — tens of mm/bp of risk vs low-single-digit-mm crossing cost → **trade.**
- **7Y split** — ~13mm to cross vs a rounding-level 7s10s residual → **skip.**
- **3M / 6M / 9M dust** — whole block ≈ 10k/bp; one tick of bid/offer ≈ 5,000× that → **skip.**
