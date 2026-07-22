"""
Business-day roll the Swap_Pricer coupon schedule.

Coupon dates were plain EDATE($B$8,12*n) with no roll, while the curve pillars
ARE modified-following rolled. Coupon dates therefore landed on weekends and
drifted off the pillars, and the interpolation error at those off-pillar dates
accumulated over 20-50 coupons - 1.3 to 2.1bp of par-rate error at 20Y+, while
1Y-10Y (few coupons) looked fine.

Effective date is set to spot, matching how an OIS is quoted and aligning the
coupon anniversaries with the curve's own pillar anniversaries.
"""
from openpyxl import load_workbook

WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg.xlsx"

def mf(e):
    nxt = f"({e}+IF(WEEKDAY({e},2)=6,2,IF(WEEKDAY({e},2)=7,1,0)))"
    prv = f"({e}-IF(WEEKDAY({e},2)=6,1,IF(WEEKDAY({e},2)=7,2,0)))"
    return f"IF(MONTH({nxt})<>MONTH({e}),{prv},{nxt})"

wb = load_workbook(WB)
sp = wb["Swap_Pricer"]
sp["B8"] = "=Bootstrap!$B$4"
sp["B10"] = "=" + mf("EDATE($B$8,12*$B$9)")
n = 0
for r in range(36, 86):
    v = sp[f"B{r}"].value
    if not (isinstance(v, str) and "EDATE" in v and "WEEKDAY" not in v):
        continue
    sp[f"B{r}"] = f'=IF(A{r}<=$B$9,{mf(f"EDATE($B$8,12*A{r})")},"")'
    n += 1
wb.calculation.fullCalcOnLoad = True
wb.save(WB)
print(f"rolled {n} coupon dates; effective=spot, maturity rolled")
