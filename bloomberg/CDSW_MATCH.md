# Matching a CDSW screen

Runbook for reproducing a CDSW capture in `CDS_Pricer.xlsx`. Written down because
the same override was typed in the wrong place twice.

## Why an override is needed at all

CDSW values on a **flat** curve at the traded spread — its Term table shows a
single row. The sheet values on the **stripped term structure**. Both are
correct; they answer different questions, and on a steep curve the difference is
real money. On the Wells Fargo 07/22/26 capture it was about 880 on 10mm.

So to reproduce a screen you feed the sheet the same flat curve CDSW is using.

## Check the quote source first

Our pull quotes `<ticker> & " BEST Curncy"` — the BEST composite. CDSW defaults
to **CMAN**, a single contributor. On Wells Fargo that alone showed 51.5600
against our 66.5597, a 15bp gap that had nothing to do with the model. Set CDSW's
CDS Curve dropdown to **BEST** before comparing anything.

## The one thing that matters

`Hazard_Bootstrap!C7 = CDS_Quotes!F7`. **Column F is the only gate.** Whatever F
says is what the strip uses. Everything below is about getting a number into F.

---

## Option A — hardcode (use this while validating)

Type the traded spread straight into `CDS_Quotes!F7:F12`, over the formulas.

One place, visible, cannot silently revert. This is the right choice for
reproducing a capture.

## Option B — override chain (for flipping between live and override)

Four hops:

```
Entities!F7  ->  Entities!F21  ->  CDS_Entities!AR5  ->  CDS_Quotes!E7  ->  F7
  input          effective         in use              manual            market
```

**Order matters. Out of sequence you lose what you have.**

### 1. Enter the spread on `Entities`

Put the traded spread in **`F7:K7`** — the 1Y / 2Y / 3Y / 5Y / 7Y / 10Y override
cells for the first entity slot. Row 7 is the first name; `F:K` are the tenors.

Do this **before** step 2. Pasting the formula first overwrites a working
hardcode with one that reads an empty `E`, falls through to BDP, and silently
restores the live curve.

### 2. Paste into `CDS_Quotes!F7`, fill down to `F12`

```excel
=IF(AND(ISNUMBER(E7),E7>0),E7,IFERROR(BDP(D7&" BEST Curncy","PX_LAST")+0,IFERROR(BDP(D7&" BEST Curncy","PX_MID")+0,0)))
```

### 3. Paste into `CDS_Quotes!H7`, fill down to `H12`

```excel
=IF(AND(ISNUMBER(E7),E7>0),"OVERRIDE",IF(ISNUMBER(IFERROR(BDP(D7&" BEST Curncy","PX_LAST")+0,"")),"BDP live","none"))
```

### 4. Verify

`H7:H12` must read **OVERRIDE** on all six. If any row says `BDP live`, the
override did not reach `E` and that tenor is still on the live pull.

---

## Also required, once

These live on sheets that do **not** travel with a moved `CDS_Pricer` tab, so
they have to be applied by hand:

`CDS_Schedule!B7`

```excel
=CDS_Parameters!$B$23
```

Period 1 must accrue from the 1st accrual start, not the valuation date. The
premium leg values the full first coupon period and (3.2) rebates the accrued
part. Truncating it leaves the premium leg short — about 4,100 on 10mm.

## What a match looks like

Wells Fargo, 07/22/26, 10mm, coupon 100bp, R 0.40, 5Y to 06/20/2031, traded
spread 66.5597 flat, BEST Mid:

| | model | CDSW | gap |
| --- | ---: | ---: | ---: |
| par spread bp | 66.5597 | 66.5597 | 0 |
| principal | -146,317 | -146,330 | 13 |
| accrued days | 31 | 31 | 0 |
| accrued | -8,611 | -8,611 | 0 |
| Prob 5Y | 0.0535 | 0.0535 | 0 |

Hazards should go **flat** at about 0.0112 across all six segments. That is the
tell that the strip is actually seeing a flat curve. If they still read
0.0043 / 0.0075 / 0.0113 / 0.0167 you are on the term structure and the override
has not landed.

## Still unverified

`CS01`, `IR DV01` and `Rec01` on `CDS_Pricer!B22:B24`. CDSW gives 4,431.93 /
35.13 / 64.15 for this trade. `CDSRisk.bas` computes all three by full revalue
with re-stripping; they have never been read off the sheet and compared. `N20`
reports whether the module is loaded.
