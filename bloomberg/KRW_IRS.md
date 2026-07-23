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
