# USD Swap Curve Bootstrap — Bloomberg Data Pull and Excel Implementation

## Objective

Construct two USD interest-rate curves:

1. **Modern USD SOFR curve** — the primary implementation.
   - SOFR OIS discount curve.
   - Built from the latest SOFR fixing, short-dated SOFR OIS instruments, SOFR futures, and longer-dated SOFR OIS swaps.
   - Intended to approximate a modern collateralised USD swap curve.

2. **Lehman-style USD LIBOR curve** — the historical research implementation.
   - Single curve used for both discounting and projecting 3M LIBOR.
   - Built from cash deposits, Eurodollar futures, and fixed-versus-3M-LIBOR swaps.
   - This is historically faithful to the 2002 methodology but is not the modern valuation standard.

The first Bloomberg session should be used to save complete, immutable market snapshots. Do not record only rates. Dates, instrument identifiers, quote timestamps, bid/ask values, conventions, and underlying curve components are also required.

---

# Part I — Data to obtain from Bloomberg now

## 1. Fix one valuation snapshot

Before pulling anything, record:

| Field | Example |
|---|---|
| Valuation date | `2026-07-20` |
| Valuation time | `16:00:00` |
| Time zone | `America/New_York` |
| Quote side | Mid |
| Bloomberg user/time stamp | Screenshot or export |
| Curve | USD OIS SOFR vs Fixed |
| Bloomberg curve number | 490 |
| Collateral assumption | USD cash collateral / SOFR discounting |
| Futures settlement or live mid | Choose one consistently |

Use the same timestamp and quote side for all instruments. If markets have closed, use official settlement prices for futures and end-of-day OIS quotes.

Save all raw exports before transforming the data.

Recommended folders:

```text
data/
├── raw/
│   └── 2026-07-20_160000_NY/
│       ├── sofr_fixing.csv
│       ├── curve_490_components.csv
│       ├── sr1_futures.csv
│       ├── sr3_futures.csv
│       ├── conventions.csv
│       └── screenshots/
└── normalized/
```

---

## 2. Bloomberg curve screen

### Terminal workflow

1. Enter:

```text
ICVS <GO>
```

2. Search for the USD SOFR OIS curve.
3. Select the curve commonly identified as:

```text
490 — USD OIS SOFR vs Fixed
```

4. Open the curve's **Components**, **Instruments**, or equivalent calibration-input view.
5. Export the complete component list to Excel.
6. Also inspect:

```text
FWCV <GO>
```

7. Enter:

```text
SOFR <GO>
```

8. Use the Bloomberg curve as a benchmark for:
   - zero rates;
   - discount factors;
   - par rates;
   - implied forwards.

Do **not** bootstrap directly from Bloomberg's already-derived zero rates if the objective is to reproduce the curve. Bootstrap from the underlying instruments and use the Bloomberg zero curve only for validation.

### Fields to export for every curve component

Export, where available:

| Required field | Purpose |
|---|---|
| Bloomberg security identifier | Reproducibility |
| Instrument description | Human verification |
| Instrument type | OIS, future, cash, etc. |
| Tenor | Ordering |
| Effective date | Schedule construction |
| Maturity date | Curve pillar |
| Bid quote | Data quality |
| Ask quote | Data quality |
| Mid or last quote | Calibration |
| Quote timestamp | Staleness check |
| Quote source/contributor | Reproducibility |
| Currency | Must be USD |
| Floating index | Must be SOFR |
| Fixed-leg frequency | Swap schedule |
| Fixed-leg day count | Swap accrual |
| Floating-leg convention | Compounded SOFR |
| Business-day convention | Date adjustment |
| Calendar | Holiday adjustment |
| Spot/effective lag | Instrument start |
| Payment lag | OIS valuation |
| Lookback/observation shift | Compounding schedule |
| Stub convention | Irregular periods |
| Curve node/pillar date | Bloomberg comparison |
| Bloomberg zero rate | Validation only |
| Bloomberg discount factor | Validation only |

If Bloomberg does not expose all conventions in the component export, open representative instruments in `SWPM <GO>` and record the conventions manually.

---

## 3. SOFR fixing

Pull the latest published SOFR fixing and sufficient historical fixings to cover any already-accrued part of:

- the first OIS coupon;
- the current 1M SOFR futures month;
- the current 3M SOFR futures reference quarter.

Recommended Bloomberg item:

```text
SOFRRATE Index
```

Pull:

| Field |
|---|
| Date |
| SOFR fixing |
| Publication date |
| Value |
| Revision flag, if available |

Excel examples:

```excel
=BDP("SOFRRATE Index","PX_LAST")
```

Historical series:

```excel
=BDH("SOFRRATE Index","PX_LAST",StartDate,EndDate)
```

Bloomberg field names and formula syntax can vary by add-in configuration. Confirm the final formula using Bloomberg's Formula Builder.

Also obtain the SOFR Index if available as a validation series. The New York Fed's published SOFR Index can independently validate the daily compounding implementation.

---

## 4. Three-month SOFR futures

### Bloomberg chain

Enter:

```text
SFRA <Comdty> CT <GO>
```

The CME's July 2026 Bloomberg code reference identifies this as the Bloomberg chain for 3-Month SOFR futures. CME Globex product code is `SR3`.

Export:

- all active serial contracts;
- all liquid quarterly IMM contracts;
- ideally through at least 3Y;
- optionally through 5Y for comparison.

For each contract obtain:

| Required field |
|---|
| Bloomberg contract ticker |
| Contract month |
| Contract description |
| Last trade date |
| Reference-period start date |
| Reference-period end date |
| Bid |
| Ask |
| Last |
| Official settlement |
| Previous settlement |
| Volume |
| Open interest |
| Quote timestamp |

Use official settlement prices for an end-of-day curve. Use mid prices only for a genuinely synchronous intraday curve.

The quoted futures rate is:

\[
R^{fut}=\frac{100-Q}{100},
\]

where \(Q\) is the futures price.

Example:

```text
Price = 95.975
Implied futures rate = (100 - 95.975)% = 4.025%
Decimal rate = 0.04025
```

Excel:

```excel
=(100-[Futures Price])/100
```

### Contract dates

For a quarterly 3M SOFR future, the reference quarter runs from the third Wednesday of the month three months before the contract month, inclusive, to the third Wednesday of the delivery month, exclusive.

Use the dates supplied by Bloomberg/CME where possible. Do not infer them solely from ticker labels.

---

## 5. One-month SOFR futures

### Bloomberg chain

Enter:

```text
SERA <Comdty> CT <GO>
```

The CME's July 2026 Bloomberg code reference identifies this as the Bloomberg chain for 1-Month SOFR futures. CME Globex product code is `SR1`.

Pull at least the nearest 6–13 monthly contracts.

Required fields are the same as for 3M SOFR futures:

- ticker;
- contract month;
- start/end dates;
- bid/ask/last/settlement;
- volume;
- open interest;
- timestamp.

For the first implementation, 1M futures are optional. They improve the front-end shape around Federal Open Market Committee dates but complicate the handling of partially accrued contracts.

Recommended first build:

```text
SOFR fixing
+ short OIS
+ 3M SOFR futures
+ long OIS swaps
```

Add 1M futures only after the basic curve works.

---

## 6. SOFR OIS quotes

From curve 490, export the exact instruments Bloomberg is using.

A practical target grid is:

```text
1W, 2W, 1M, 2M, 3M, 4M, 5M, 6M, 9M,
1Y, 18M, 2Y, 3Y, 4Y, 5Y, 6Y, 7Y, 8Y,
9Y, 10Y, 12Y, 15Y, 20Y, 25Y, 30Y, 40Y, 50Y
```

Do not force this grid if Bloomberg curve 490 uses a different set. The exact exported component set is authoritative for the Bloomberg-replication exercise.

Common Bloomberg SOFR swap identifiers often follow a pattern such as:

```text
USOSFR1 Curncy
USOSFR2 Curncy
USOSFR3 Curncy
USOSFR5 Curncy
...
```

However, do not rely on a guessed ticker pattern. Copy the exact identifiers from curve 490's component list because Bloomberg instrument identifiers, source settings, and entitlements can differ.

For each OIS quote, save:

| Field |
|---|
| Exact Bloomberg security |
| Tenor |
| Par fixed rate |
| Bid |
| Ask |
| Mid |
| Effective date |
| Maturity date |
| Fixed payment dates |
| Fixed accrual fractions |
| Fixed frequency |
| Fixed day count |
| Floating index |
| Floating compounding convention |
| Payment delay |
| Calendar |
| Business-day adjustment |
| Timestamp |

---

## 7. Bloomberg validation outputs

Export the following from curve 490 or `FWCV` for comparison after bootstrapping:

| Output |
|---|
| Pillar dates |
| Discount factors |
| Continuously compounded zero rates |
| Annually compounded zero rates |
| 1-day or instantaneous forwards |
| 1M forward SOFR |
| 3M-equivalent forward SOFR |
| Par OIS rates |
| Curve metadata/interpolation setting |

These are validation targets, not calibration inputs.

---



# Part I-A — Bloomberg Excel commands: copy-and-paste tables

## 1. Bloomberg Excel function map

| Function | Use in this project | Shape of result | Example |
|---|---|---|---|
| `BDP` | Current/reference data for one security and one or more fields | One value per field | `=BDP("SOFRRATE Index","PX_LAST")` |
| `BDH` | Historical time series between two dates | Date/value table | `=BDH("SOFRRATE Index","PX_LAST",$B$2,$B$3)` |
| `BDS` | Bulk datasets such as a futures chain | Spilled table | `=BDS("SFRA Comdty","FUT_CHAIN")` |
| `BDP` with field array | Current snapshot for several fields on one instrument | Horizontal/vertical array depending on Excel setup | `=BDP($A2,$B$1:$K$1)` |
| `BDH` with field array | Historical values for multiple fields | Historical table | `=BDH($A2,$B$1:$D$1,$B$2,$B$3)` |

Bloomberg formulas can be generated and checked through:

```text
Bloomberg Excel Add-In → Formula Builder
```

Use Bloomberg's field search if a field below is unavailable under your entitlement. On the Terminal, enter:

```text
FLDS <GO>
```

The syntax in this section uses commas. Some regional Excel installations use semicolons instead.

---

## 2. Control cells used by the formulas

Create these cells on the `Control` sheet:

| Cell | Name | Example |
|---|---|---|
| `B2` | `ValuationDate` | `20-Jul-2026` |
| `B3` | `HistoryStart` | `01-Jan-2026` |
| `B4` | `HistoryEnd` | `20-Jul-2026` |
| `B5` | `PeriodicSelection` | `DAILY` |
| `B6` | `QuoteSide` | `MID` |

Optionally assign Excel names to these cells. The examples below use direct cell references so they can be pasted immediately.

---

## 3. SOFR fixing — current and historical

Paste this table into Excel if desired:

```text
Purpose	Security	Bloomberg Excel formula
Latest SOFR fixing	SOFRRATE Index	=BDP("SOFRRATE Index","PX_LAST")
Latest fixing date	SOFRRATE Index	=BDP("SOFRRATE Index","REFERENCE_DATE")
SOFR name/description	SOFRRATE Index	=BDP("SOFRRATE Index","NAME")
Historical SOFR fixing	SOFRRATE Index	=BDH("SOFRRATE Index","PX_LAST",Control!$B$3,Control!$B$4,"Per=DAILY","Days=T","Dir=V","Dts=S")
Historical fixing and publication metadata	SOFRRATE Index	=BDH("SOFRRATE Index",{"PX_LAST","REFERENCE_DATE"},Control!$B$3,Control!$B$4,"Per=DAILY","Days=T","Dir=V","Dts=S")
```

Recommended single-cell formulas:

```excel
=BDP("SOFRRATE Index","PX_LAST")
```

```excel
=BDH(
    "SOFRRATE Index",
    "PX_LAST",
    Control!$B$3,
    Control!$B$4,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

Notes:

- `Days=T` requests trading-day observations rather than filling every calendar day.
- Do not use filled-forward weekend values from Bloomberg as separate fixings. In the compounding sheet, apply the preceding business-day SOFR across the relevant number of calendar days.
- `REFERENCE_DATE` availability can depend on the security and entitlement. Confirm in `FLDS <GO>`.

---

## 4. SOFR futures chains

### Terminal commands

```text
SFRA <Comdty> CT <GO>    3M SOFR futures chain
SERA <Comdty> CT <GO>    1M SOFR futures chain
```

### Excel bulk-chain formulas

```text
Purpose	Security	Bloomberg Excel formula
3M SOFR futures chain	SFRA Comdty	=BDS("SFRA Comdty","FUT_CHAIN")
1M SOFR futures chain	SERA Comdty	=BDS("SERA Comdty","FUT_CHAIN")
```

Paste into separate worksheets because `BDS` returns a spilled table:

```excel
=BDS("SFRA Comdty","FUT_CHAIN")
```

```excel
=BDS("SERA Comdty","FUT_CHAIN")
```

Depending on the Bloomberg Excel version, `FUT_CHAIN` may return one identifier column or multiple descriptive columns. Keep the returned chain as the canonical list of active contracts.

### Pull fields for each futures contract

Assume the chain tickers are in `Futures!A2:A50`. Put these headers in row 1:

```text
A1  Security
B1  NAME
C1  PX_BID
D1  PX_ASK
E1  PX_LAST
F1  PX_SETTLE
G1  FUT_CONTRACT_DATE
H1  LAST_TRADEABLE_DT
I1  FUT_NOTICE_FIRST
J1  FUT_DLV_DT_FIRST
K1  FUT_DLV_DT_LAST
L1  VOLUME
M1  OPEN_INT
N1  CRNCY
O1  EXCH_CODE
P1  SECURITY_TYP
Q1  TIME_LAST_UPDATE
R1  DATE
```

In `B2`, use a field-range pull and copy down:

```excel
=BDP($A2,B$1:R$1)
```

If the Excel add-in does not accept a horizontal field range in one formula, use individual formulas:

```text
Column	Formula to paste in row 2
NAME	=BDP($A2,"NAME")
Bid	=BDP($A2,"PX_BID")
Ask	=BDP($A2,"PX_ASK")
Last	=BDP($A2,"PX_LAST")
Settlement	=BDP($A2,"PX_SETTLE")
Contract month	=BDP($A2,"FUT_CONTRACT_DATE")
Last trade date	=BDP($A2,"LAST_TRADEABLE_DT")
First notice date	=BDP($A2,"FUT_NOTICE_FIRST")
First delivery date	=BDP($A2,"FUT_DLV_DT_FIRST")
Last delivery date	=BDP($A2,"FUT_DLV_DT_LAST")
Volume	=BDP($A2,"VOLUME")
Open interest	=BDP($A2,"OPEN_INT")
Currency	=BDP($A2,"CRNCY")
Exchange	=BDP($A2,"EXCH_CODE")
Security type	=BDP($A2,"SECURITY_TYP")
Last update time	=BDP($A2,"TIME_LAST_UPDATE")
Bloomberg date	=BDP($A2,"DATE")
```

Not every Bloomberg futures security exposes every candidate field. Use `FLDS <GO>` on a representative SR1/SR3 contract and substitute the Bloomberg field that your terminal returns. In particular, contract reference-period dates may need to be obtained from the contract description or CME specification if Bloomberg exposes only delivery/last-trade dates.

### Historical futures settlements

Assume the futures ticker is in `A2`:

```excel
=BDH(
    $A2,
    "PX_SETTLE",
    Control!$B$3,
    Control!$B$4,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

For several fields:

```excel
=BDH(
    $A2,
    {"PX_SETTLE","PX_LAST","VOLUME","OPEN_INT"},
    Control!$B$3,
    Control!$B$4,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

### Current futures quote selection

If bid and ask exist:

```excel
=AVERAGE(C2,D2)
```

Robust fallback:

```excel
=IF(
    AND(ISNUMBER(C2),ISNUMBER(D2)),
    (C2+D2)/2,
    IF(ISNUMBER(F2),F2,E2)
)
```

Implied futures rate in decimal:

```excel
=(100-SelectedFuturesPrice)/100
```

---

## 5. SOFR OIS curve instruments

### Preferred method: export the Bloomberg curve component list

On the Terminal:

```text
ICVS <GO>
```

Open USD curve 490 or the current Bloomberg USD SOFR OIS curve, then export its component securities to Excel. Paste the exact component identifiers into `OIS_Quotes!A2:A100`.

Do not manufacture the OIS ticker list from a presumed naming convention. The curve component export is preferable because it captures the exact Bloomberg securities and sources used by that curve.

### Headers for OIS quote pulls

Put these fields in row 1:

```text
A1  Security
B1  NAME
C1  PX_BID
D1  PX_ASK
E1  PX_LAST
F1  PX_MID
G1  MATURITY
H1  DAYS_TO_MTY
I1  CRNCY
J1  SECURITY_TYP
K1  DAY_CNT_DES
L1  CPN_FREQ
M1  SETTLE_DT
N1  TIME_LAST_UPDATE
O1  DATE
P1  ID_BB_GLOBAL
```

In `B2`, attempt:

```excel
=BDP($A2,B$1:P$1)
```

Or use separate formulas:

```text
Column	Formula to paste in row 2
Description	=BDP($A2,"NAME")
Bid	=BDP($A2,"PX_BID")
Ask	=BDP($A2,"PX_ASK")
Last	=BDP($A2,"PX_LAST")
Mid	=BDP($A2,"PX_MID")
Maturity	=BDP($A2,"MATURITY")
Days to maturity	=BDP($A2,"DAYS_TO_MTY")
Currency	=BDP($A2,"CRNCY")
Security type	=BDP($A2,"SECURITY_TYP")
Day-count description	=BDP($A2,"DAY_CNT_DES")
Coupon frequency	=BDP($A2,"CPN_FREQ")
Settlement date	=BDP($A2,"SETTLE_DT")
Last update time	=BDP($A2,"TIME_LAST_UPDATE")
Bloomberg date	=BDP($A2,"DATE")
Bloomberg global ID	=BDP($A2,"ID_BB_GLOBAL")
```

Some OTC swap curve tickers will not populate bond-style fields such as `MATURITY`, `CPN_FREQ`, or `SETTLE_DT`. In that case:

1. retain the exact Bloomberg component ticker;
2. record the tenor from the component screen;
3. open the representative swap in `SWPM <GO>`;
4. manually capture effective lag, payment frequency, day count, calendar, business-day rule, payment lag, and observation convention.

### Historical OIS quote pull

Assume the curve component ticker is in `A2`:

```excel
=BDH(
    $A2,
    {"PX_BID","PX_ASK","PX_LAST"},
    Control!$B$3,
    Control!$B$4,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

For a single valuation date, use the same date as start and end:

```excel
=BDH(
    $A2,
    {"PX_BID","PX_ASK","PX_LAST"},
    Control!$B$2,
    Control!$B$2,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

If the valuation date is a holiday or a non-quoted date, explicitly choose whether to use the previous valid business date. Do not allow Excel to silently substitute a value without documenting the decision.

---

## 6. Bloomberg-derived curve outputs for validation

The preferred method is to export curve outputs from `ICVS <GO>` or `FWCV <GO>` directly. If Bloomberg provides explicit securities or fields for the selected curve, use Formula Builder to generate those formulas rather than guessing field identifiers.

Terminal commands:

```text
ICVS <GO>    curve components and curve settings
FWCV <GO>    forward curve output
SWPM <GO>    swap conventions and repricing
FLDS <GO>    field lookup
DES  <GO>    security description and metadata
HP   <GO>    historical price table
GP   <GO>    historical chart
```

Export the following validation columns:

```text
Pillar date
Discount factor
Zero rate
Forward rate
Par OIS rate
Interpolation method
Curve number/name
Valuation timestamp
```

These are comparison outputs. Do not feed Bloomberg discount factors or zero rates into the instrument bootstrap if the objective is to reproduce the curve independently.

---

## 7. Generic copy-paste quote template

Use this when a list of Bloomberg securities has already been pasted into column `A`.

```text
Security	Name	Bid	Ask	Last	Settlement	Mid/Fallback	Timestamp	Date
<PASTE TICKER>	=BDP($A2,"NAME")	=BDP($A2,"PX_BID")	=BDP($A2,"PX_ASK")	=BDP($A2,"PX_LAST")	=BDP($A2,"PX_SETTLE")	=IF(AND(ISNUMBER(C2),ISNUMBER(D2)),(C2+D2)/2,IF(ISNUMBER(F2),F2,E2))	=BDP($A2,"TIME_LAST_UPDATE")	=BDP($A2,"DATE")
```

Copy the formula row down for all securities.

---

## 8. Historical Lehman-era Bloomberg pulls

The exact historical security identifiers should be found on the Terminal and pasted into a mapping table. Use `SECF <GO>`, security search, or curve component screens rather than assuming that all discontinued tickers retain current naming patterns.

### Historical fixing/rate template

```excel
=BDH(
    $A2,
    "PX_LAST",
    Control!$B$2,
    Control!$B$2,
    "Per=DAILY",
    "Days=T",
    "Dir=V",
    "Dts=S"
)
```

Suggested mapping table:

```text
Instrument	Bloomberg security found on Terminal	Field	Valuation date	Formula
1M USD LIBOR	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B2,C2,D2,D2,"Per=DAILY","Days=T","Dir=V","Dts=S")
3M USD LIBOR	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B3,C3,D3,D3,"Per=DAILY","Days=T","Dir=V","Dts=S")
6M USD LIBOR	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B4,C4,D4,D4,"Per=DAILY","Days=T","Dir=V","Dts=S")
12M USD LIBOR	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B5,C5,D5,D5,"Per=DAILY","Days=T","Dir=V","Dts=S")
Eurodollar future 1	<PASTE>	PX_SETTLE	26-Aug-2002	=BDH(B6,C6,D6,D6,"Per=DAILY","Days=T","Dir=V","Dts=S")
2Y USD LIBOR swap	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B7,C7,D7,D7,"Per=DAILY","Days=T","Dir=V","Dts=S")
5Y USD LIBOR swap	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B8,C8,D8,D8,"Per=DAILY","Days=T","Dir=V","Dts=S")
10Y USD LIBOR swap	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B9,C9,D9,D9,"Per=DAILY","Days=T","Dir=V","Dts=S")
30Y USD LIBOR swap	<PASTE>	PX_LAST	26-Aug-2002	=BDH(B10,C10,D10,D10,"Per=DAILY","Days=T","Dir=V","Dts=S")
```

Because LIBOR and Eurodollar instruments are discontinued, historical-field availability and ticker continuity must be checked directly on Bloomberg. If Bloomberg returns no observation, record the failure rather than substituting a modern instrument.

---

## 9. Bloomberg formula error handling

Common Bloomberg Excel errors:

| Error | Likely cause | Action |
|---|---|---|
| `#N/A Invalid Security` | Wrong or incomplete ticker | Confirm full yellow-key identifier on Terminal |
| `#N/A Invalid Field` | Field not supported for that security | Use `FLDS <GO>` or Formula Builder |
| `#N/A N/A` | No value or entitlement | Check Terminal screen and entitlement |
| `#N/A Requesting Data...` | Add-in is refreshing | Wait for synchronous refresh before saving snapshot |
| Blank historical result | Non-business date or unavailable history | Pull a wider date window and select the intended observation explicitly |
| Stale quote | OTC component has not updated | Record timestamp and decide whether to exclude it |

Before freezing the snapshot:

1. refresh the workbook;
2. wait until no cells show `Requesting Data`;
3. copy the entire raw workbook;
4. paste values into a dated immutable snapshot workbook;
5. retain the live Bloomberg-linked workbook separately;
6. record the refresh timestamp and time zone.

---

## 10. Minimal formula set required for the first build

```text
Data item	Bloomberg Excel formula
Latest SOFR	=BDP("SOFRRATE Index","PX_LAST")
SOFR history	=BDH("SOFRRATE Index","PX_LAST",Control!$B$3,Control!$B$4,"Per=DAILY","Days=T","Dir=V","Dts=S")
3M SOFR chain	=BDS("SFRA Comdty","FUT_CHAIN")
1M SOFR chain	=BDS("SERA Comdty","FUT_CHAIN")
Futures bid	=BDP($A2,"PX_BID")
Futures ask	=BDP($A2,"PX_ASK")
Futures settlement	=BDP($A2,"PX_SETTLE")
Futures last trade date	=BDP($A2,"LAST_TRADEABLE_DT")
Futures volume	=BDP($A2,"VOLUME")
Futures open interest	=BDP($A2,"OPEN_INT")
OIS bid	=BDP($A2,"PX_BID")
OIS ask	=BDP($A2,"PX_ASK")
OIS last	=BDP($A2,"PX_LAST")
OIS historical snapshot	=BDH($A2,{"PX_BID","PX_ASK","PX_LAST"},Control!$B$2,Control!$B$2,"Per=DAILY","Days=T","Dir=V","Dts=S")
```

The exact curve-component securities must still be exported from Bloomberg curve 490 and pasted into column `A`. This is preferable to hard-coding a potentially incorrect OIS ticker list.


# Part II — Recommended calibration set

## Version 1: simplest robust modern curve

Use:

| Segment | Instruments |
|---|---|
| Anchor | Valuation-date discount factor \(P(0,0)=1\) |
| Overnight | Latest SOFR fixing |
| Short end | SOFR OIS through 12M |
| Middle | 3M SOFR futures from first clean future contract through 3Y |
| Long end | SOFR OIS swaps from 4Y through final available maturity |
| Interpolation | Linear interpolation in log discount factors |
| Futures convexity | Zero initially |
| Quote side | Mid or settlement, consistently |

A **clean future contract** is one whose reference period has not started. Partially accrued contracts require realised SOFR fixings and separate treatment.

## Version 2: improved front end

Use:

- short OIS;
- 1M futures for near-month/FOMC granularity;
- 3M futures thereafter;
- long OIS swaps.

## Version 3: closer Bloomberg replication

Use the exact Bloomberg curve-490 components and conventions, including:

- exact instrument cutoffs;
- exact interpolation;
- exact convexity handling;
- exact payment lag;
- exact stub rules;
- exact quote source.

---

# Part III — Excel workbook layout

Create these worksheets:

```text
1. Control
2. Raw_BBG
3. Fixings
4. Instruments
5. Futures
6. Curve_Nodes
7. OIS_Schedules
8. Calibration
9. Outputs
10. Bloomberg_Compare
```

---

## Sheet 1 — `Control`

Suggested cells:

| Cell | Value |
|---|---|
| B2 | Valuation date |
| B3 | Valuation timestamp |
| B4 | Quote side |
| B5 | Curve name |
| B6 | Futures cutoff |
| B7 | OIS restart tenor |
| B8 | Convexity mode |
| B9 | Interpolation |
| B10 | Root-solving tolerance |

Example:

```text
B2 = 20-Jul-2026
B4 = MID
B5 = USD_SOFR_OIS
B6 = 3Y
B7 = 4Y
B8 = ZERO
B9 = LOG_LINEAR_DF
B10 = 0.0000000001
```

---

## Sheet 2 — `Raw_BBG`

Paste Bloomberg outputs without modifying them.

Columns:

```text
A  Snapshot timestamp
B  Bloomberg security
C  Instrument type
D  Tenor/contract
E  Bid
F  Ask
G  Last
H  Settlement
I  Quote timestamp
J  Effective date
K  Maturity/reference end
L  Volume
M  Open interest
N  Source
O  Notes
```

Mid quote:

```excel
=IF(AND(E2<>"",F2<>""),(E2+F2)/2,IF(H2<>"",H2,G2))
```

Never overwrite the raw values.

---

## Sheet 3 — `Fixings`

Columns:

```text
A  SOFR date
B  SOFR percent
C  SOFR decimal
D  Next SOFR publication/business date
E  Calendar days applying
F  Daily accrual factor
G  Daily growth factor
H  Cumulative index
```

Formulas:

```excel
C2 = B2/100
E2 = A3-A2
F2 = E2/360
G2 = 1+C2*F2
H2 = 1
H3 = H2*G2
```

The New York Fed methodology applies the previous business day's SOFR across intervening non-business days. Consequently, `E2` may equal 1, 3, or another number depending on weekends and holidays.

Compounded SOFR over dates \(a\) to \(b\):

\[
R_{a,b}
=
\frac{360}{d_c}
\left(
\frac{I_b}{I_a}-1
\right),
\]

where \(d_c\) is the number of calendar days and \(I\) is the compounded index.

Excel:

```excel
=360/(EndDate-StartDate)*(EndIndex/StartIndex-1)
```

---

## Sheet 4 — `Instruments`

Normalize all calibration instruments.

Columns:

```text
A  Bootstrap order
B  Instrument ID
C  Type
D  Tenor
E  Start date
F  End/pillar date
G  Market quote
H  Quote type
I  Day-count fraction
J  Convexity adjustment
K  Adjusted rate
L  Include?
M  Data-quality flag
```

For futures:

```excel
K2 = (100-G2)/100-J2
```

For OIS:

```excel
K2 = G2/100
```

if Bloomberg exports the quote in percentage points.

Check units carefully. A Bloomberg quote of `4.025` normally means 4.025%, which must be represented as `0.04025` in pricing formulas.

---

# Part IV — Core curve representation

## 1. Discount factors

The curve is represented by discount factors:

\[
P(0,T).
\]

Set:

\[
P(0,0)=1.
\]

## 2. Year fractions

For the basic Excel version, use:

```excel
=(Date-ValuationDate)/365
```

only for plotting zero rates.

Do not use this generic year fraction for pricing cash flows. Use each instrument's contractual day-count convention.

## 3. Zero rates

Continuously compounded zero rate:

\[
z_c(T)=-\frac{\ln P(0,T)}{T}.
\]

Excel:

```excel
=-LN(DiscountFactor)/YearFraction
```

Annually compounded zero rate:

\[
z_a(T)=P(0,T)^{-1/T}-1.
\]

Excel:

```excel
=DiscountFactor^(-1/YearFraction)-1
```

## 4. Forward rates

Simple forward rate over \([T_1,T_2]\):

\[
F(T_1,T_2)
=
\frac{1}{\delta}
\left[
\frac{P(0,T_1)}{P(0,T_2)}-1
\right].
\]

Excel:

```excel
=(DF_Start/DF_End-1)/AccrualFraction
```

---

# Part V — Log-linear discount-factor interpolation in Excel

Suppose:

- known node 1 is at date \(T_1\) with discount factor \(P_1\);
- known node 2 is at date \(T_2\) with discount factor \(P_2\);
- the required cash-flow date is \(t\).

Define:

\[
w=\frac{t-T_1}{T_2-T_1}.
\]

Then:

\[
\ln P(t)
=
(1-w)\ln P_1+w\ln P_2.
\]

Therefore:

\[
P(t)
=
\exp\left[(1-w)\ln P_1+w\ln P_2\right].
\]

Excel:

```excel
=EXP(
    (1-(TargetDate-LeftDate)/(RightDate-LeftDate))*LN(LeftDF)
    +((TargetDate-LeftDate)/(RightDate-LeftDate))*LN(RightDF)
)
```

Equivalent compact formula:

```excel
=LeftDF^(1-w)*RightDF^w
```

Do not interpolate par swap rates directly.

---

# Part VI — Bootstrap the short OIS segment

## 1. Simplified OIS identity

For a spot-starting par OIS with fixed rate \(K\), fixed payment dates \(T_1,\ldots,T_n\), and accrual factors \(\alpha_i\):

\[
K\sum_{i=1}^{n}\alpha_iP(0,T_i)
=
P(0,T_0)-P(0,T_n).
\]

For a spot-starting instrument with \(P(0,T_0)\) known:

\[
P(0,T_n)
=
\frac{
P(0,T_0)
-
K\sum_{i=1}^{n-1}\alpha_iP(0,T_i)
}{
1+K\alpha_n
}.
\]

If the OIS has only one payment at maturity:

\[
P(0,T)
=
\frac{P(0,T_0)}{1+K\alpha}.
\]

Excel for a one-payment OIS:

```excel
=StartDF/(1+Rate*AccrualFraction)
```

Excel for a multi-payment OIS:

```excel
=(StartDF-Rate*SUMPRODUCT(PreviousAccruals,PreviousDFs))
 /(1+Rate*FinalAccrual)
```

### Important limitation

The closed-form formula works cleanly only where all earlier coupon-date discount factors are already known or interpolated independently of the new terminal node.

If one or more intermediate coupon dates depend on the new node through interpolation, use Excel Solver or Goal Seek.

---

# Part VII — Bootstrap SOFR futures

## 1. Convert price into rate

\[
R_i^{fut}=\frac{100-Q_i}{100}.
\]

Excel:

```excel
=(100-Price)/100
```

## 2. Apply convexity adjustment

Version 1:

\[
CA_i=0.
\]

Then:

\[
F_i=R_i^{fut}.
\]

Improved version:

\[
F_i=R_i^{fut}-CA_i.
\]

Keep convexity in a separate input column so it can be changed without rewriting formulas.

## 3. Bootstrap the end discount factor

For a clean futures reference period \([T_i,T_{i+1}]\):

\[
P(0,T_{i+1})
=
\frac{P(0,T_i)}
{1+\delta_iF_i}.
\]

Excel:

```excel
=StartDF/(1+AccrualFraction*AdjustedForwardRate)
```

For 3M SOFR futures, use the actual reference-period day count:

```excel
=(ReferenceEnd-ReferenceStart)/360
```

subject to the exact contract methodology and date conventions.

## 4. Start-date interpolation

The first future's start date may fall between existing OIS nodes.

Interpolate \(P(0,T_i)\) using log-linear discount factors, then calculate the end discount factor.

## 5. Partially accrued futures

If the reference period has started, divide the total compounded return into:

1. realised SOFR growth;
2. unknown future growth.

Let:

\[
G_{\mathrm{total}}
=
1+R^{fut}\Delta.
\]

Let realised growth be:

\[
G_{\mathrm{realised}}
=
\prod_{k\in\text{realised}}
(1+r_kd_k).
\]

Then required future growth is:

\[
G_{\mathrm{future}}
=
\frac{G_{\mathrm{total}}}{G_{\mathrm{realised}}}.
\]

The implied forward rate over the remaining period is:

\[
F_{\mathrm{remaining}}
=
\frac{G_{\mathrm{future}}-1}{\Delta_{\mathrm{remaining}}}.
\]

Excel:

```excel
TotalGrowth     = 1+FuturesRate*TotalAccrual
RealisedGrowth  = PRODUCT(1+DailySOFR*DailyAccrual)
FutureGrowth    = TotalGrowth/RealisedGrowth
RemainingRate   = (FutureGrowth-1)/RemainingAccrual
```

For the initial build, exclude partially accrued futures and begin with the first clean contract.

---

# Part VIII — Bootstrap longer SOFR OIS swaps

## 1. Build the fixed schedule

For each OIS tenor:

1. determine effective date;
2. generate fixed payment dates;
3. adjust dates using the correct calendar and business-day convention;
4. calculate fixed accrual factors;
5. determine payment dates including payment lag;
6. obtain/interpolate a discount factor for each payment date.

Suggested `OIS_Schedules` columns:

```text
A  Swap ID
B  Coupon number
C  Accrual start
D  Accrual end
E  Payment date
F  Accrual fraction
G  Discount factor
H  Discounted accrual
```

Formula:

```excel
H2 = F2*G2
```

Fixed-leg annuity:

\[
A=\sum_i\alpha_iP(0,T_i).
\]

Excel:

```excel
=SUM(H:H)
```

Fixed-leg present value per unit notional:

\[
PV_{\mathrm{fixed}}=K A.
\]

## 2. Floating leg

For a spot-starting OIS under the standard telescoping representation:

\[
PV_{\mathrm{float}}
=
P(0,T_0)-P(0,T_n).
\]

## 3. Par equation

\[
NPV
=
K\sum_i\alpha_iP(0,T_i)
-
\left[P(0,T_0)-P(0,T_n)\right].
\]

At calibration:

\[
NPV=0.
\]

Excel model quote:

```excel
=(StartDF-EndDF)/SUMPRODUCT(Accruals,PaymentDFs)
```

Calibration error:

```excel
=ModelParRate-MarketParRate
```

## 4. Use Goal Seek

For each long OIS swap:

1. enter an initial guess for the new terminal discount factor;
2. use that discount factor as the new curve node;
3. interpolate all intermediate payment-date discount factors;
4. calculate the model par rate;
5. calculate `Model Rate - Market Rate`;
6. select:

```text
Data → What-If Analysis → Goal Seek
```

7. Set:
   - **Set cell:** calibration-error cell;
   - **To value:** `0`;
   - **By changing cell:** terminal discount-factor cell.

Repeat in maturity order.

## 5. Use Solver for the full curve

A more scalable alternative is to solve all unknown discount factors simultaneously.

Objective:

\[
\min \sum_i
\left(
K_i^{model}-K_i^{market}
\right)^2.
\]

Excel objective cell:

```excel
=SUMSQ(CalibrationErrorRange)
```

Solver setup:

```text
Set Objective: TotalSquaredError
To: Min
By Changing: UnknownDiscountFactorRange
Constraints:
    Discount factors > 0
Optional:
    DF(i+1) <= DF(i), where appropriate
Method:
    GRG Nonlinear
```

Sequential Goal Seek is easier to audit. Solver is easier where instruments overlap or payment dates depend on interpolation involving several unknown nodes.

---

# Part IX — Recommended Excel bootstrap order

Use this exact order for Version 1:

```text
Step 1  Set P(0, valuation date) = 1
Step 2  Load SOFR fixing history
Step 3  Bootstrap short OIS instruments through 12M
Step 4  Interpolate to the first clean 3M future start date
Step 5  Bootstrap consecutive 3M futures through 3Y
Step 6  Bootstrap 4Y OIS
Step 7  Bootstrap 5Y OIS
Step 8  Continue through 50Y
Step 9  Reprice every calibration instrument
Step 10 Compare against Bloomberg curve 490
```

Do not calibrate both a future and an OIS as exact nodes at the same maturity in the first version. Establish an explicit cutoff.

---

# Part X — Calibration report

Create one row per instrument:

```text
A  Instrument
B  Type
C  Market quote
D  Model quote
E  Error decimal
F  Error basis points
G  NPV error
H  Pillar date
I  Included?
J  Comment
```

Basis-point error:

```excel
=(ModelQuote-MarketQuote)*10000
```

Pass criteria for the first implementation:

```text
Synthetic instruments: less than 1e-8 rate error
Real Bloomberg rounded quotes: less than 0.01 bp where exact fitting is intended
```

If the quote itself is rounded, exact machine-precision matching to Bloomberg's derived curve may be impossible.

---

# Part XI — Output sheet

For a regular daily, monthly, or tenor grid, calculate:

| Output |
|---|
| Date |
| Year fraction |
| Discount factor |
| Continuous zero rate |
| Annual zero rate |
| 1D forward |
| 1M forward |
| 3M forward |
| Par OIS rate |

Example formulas:

```excel
ContinuousZero = -LN(DF)/T
AnnualZero     = DF^(-1/T)-1
Forward        = (DF_Start/DF_End-1)/Accrual
```

Plot:

1. par OIS input rates;
2. zero curve;
3. forward curve;
4. Bloomberg-minus-model zero-rate difference;
5. Bloomberg-minus-model par-rate difference.

---

# Part XII — Bloomberg comparison

Suggested columns:

```text
A  Date/tenor
B  Model DF
C  Bloomberg DF
D  DF difference
E  Model zero
F  Bloomberg zero
G  Zero difference in bp
H  Model par rate
I  Bloomberg par rate
J  Par difference in bp
```

Formula:

```excel
=(ModelZero-BloombergZero)*10000
```

Investigate differences in this order:

1. quote timestamp;
2. instrument list;
3. bid/mid/ask selection;
4. futures settlement versus live price;
5. spot/effective date;
6. holiday calendar;
7. fixed day count;
8. payment frequency;
9. payment lag;
10. observation shift/lookback;
11. futures convexity;
12. interpolation variable;
13. interpolation method;
14. curve-pillar selection;
15. extrapolation.

Do not compensate for unexplained differences by manually shifting the final curve.

---

# Part XIII — Historical Lehman-style Bloomberg pull

If the historical 2002 curve is also required while Bloomberg access remains available, save:

## Cash/short LIBOR

```text
1M USD LIBOR
3M USD LIBOR
6M USD LIBOR
12M USD LIBOR
```

Use historical values for the selected 2002 valuation date.

## Eurodollar futures

Save the full active Eurodollar futures strip for the historical date, including:

- exact contract identifiers;
- prices;
- IMM start/end dates;
- settlement;
- volume/open interest if available.

## USD LIBOR swaps

Save fixed-versus-3M-LIBOR par swap rates, ideally:

```text
2Y, 3Y, 4Y, 5Y, 7Y, 10Y, 12Y, 15Y, 20Y, 25Y, 30Y
```

## Historical conventions

Record:

- deposit day count;
- spot lag;
- fixed-leg frequency;
- fixed-leg day count;
- 3M LIBOR reset convention;
- business-day convention;
- calendar;
- futures convexity methodology.

The historical curve uses one discount/projection curve:

\[
F^{3M}(T_i,T_{i+1})
=
\frac{1}{\delta_i}
\left(
\frac{P(0,T_i)}{P(0,T_{i+1})}-1
\right).
\]

The same discount factors are used to discount both fixed and floating cash flows.

---

# Part XIV — Minimum Bloomberg checklist before losing access

Save these items now:

- [ ] Curve 490 complete component export
- [ ] Curve 490 screenshots showing settings and interpolation
- [ ] Exact Bloomberg identifiers for every OIS component
- [ ] Bid, ask, mid, and timestamps
- [ ] SOFR fixing history
- [ ] SOFR Index history, if available
- [ ] `SFRA <Comdty> CT` complete 3M futures chain
- [ ] `SERA <Comdty> CT` complete 1M futures chain
- [ ] Futures contract reference dates
- [ ] Futures settlements, volume, and open interest
- [ ] SWPM convention screenshots for representative OIS maturities
- [ ] Bloomberg discount factors
- [ ] Bloomberg zero rates
- [ ] Bloomberg implied forward rates
- [ ] Bloomberg par-rate outputs
- [ ] Any Bloomberg convexity settings
- [ ] Bloomberg curve interpolation and extrapolation settings
- [ ] Historical 2002 LIBOR, Eurodollar futures, and swap snapshots if accessible
- [ ] All raw exports saved without transformations
- [ ] Retrieval timestamp and terminal user recorded

---

# Part XV — Source notes

1. Bloomberg identifies `FWCV <GO>` as its forward-curve function and states that the displayed spot and projection columns are derived from bootstrapped interest-rate swap quotations.
2. The CME identifies the Bloomberg 3M SOFR futures chain as `SFRA <Comdty> CT`, with CME Globex code `SR3`.
3. The CME identifies the Bloomberg 1M SOFR futures chain as `SERA <Comdty> CT`, with CME Globex code `SR1`.
4. Three-month SOFR futures settle against business-day compounded SOFR over the contract reference quarter and are quoted as 100 minus the annualised compounded rate.
5. The New York Fed's SOFR Averages and Index use daily compounding on business days and apply the preceding business day's SOFR over non-business days.

Useful references:

- Bloomberg forward-curve overview:  
  https://www.bloomberg.com/professional/insights/treasury/forecasting-interest-rate-expenses-in-a-volatile-market/

- CME Bloomberg code reference:  
  https://www.cmegroup.com/articles/2024/bloomberg-codes-reference-sheet.html

- CME 3M SOFR contract specifications:  
  https://www.cmegroup.com/markets/interest-rates/stirs/three-month-sofr.contractSpecs.html

- New York Fed SOFR data:  
  https://www.newyorkfed.org/markets/reference-rates/sofr

- New York Fed SOFR Averages and Index:  
  https://www.newyorkfed.org/markets/reference-rates/sofr-averages-and-index

- SOFR bootstrapping paper with Python/Excel discussion:  
  https://papers.ssrn.com/sol3/papers.cfm?abstract_id=3654466

---

# Initial implementation recommendation

Build the first Excel curve using:

```text
Curve:             USD SOFR OIS
Short OIS:         Through 12M
Futures:           First clean 3M SOFR future through 3Y
Long OIS:          4Y through final Bloomberg curve-490 maturity
Interpolation:     Log-linear discount factors
Futures convexity: Zero
Quote side:        Mid or official settlement
Solver:            Sequential Goal Seek
```

Once this calibrates and reprices correctly:

1. add partially accrued futures;
2. add 1M SOFR futures;
3. add convexity adjustments;
4. reproduce Bloomberg's exact interpolation;
5. compare the resulting curve with curve 490;
6. implement the same logic in Python.
