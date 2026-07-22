"""
F22 (the first annual coupon accrual) still ran from SPOT while the rest of the
bootstrap discounts from SETTLE. That 2-day mismatch enters the annuity of every
long-end pillar and compounds: the workbook read max |dz| 1.485bp against an
independently verified 0.36bp.

Also point Swap_Pricer's effective date at the settle date so a par swap mirrors
the bootstrap instrument exactly, and business-day roll its maturity so it lands
on the same pillar date the curve was built from.
"""
from openpyxl import load_workbook

WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg_Pricer.xlsx"
wb = load_workbook(WB)
wb["Bootstrap"]["F22"] = "=(B22-VAL_DATE)/360"

sp = wb["Swap_Pricer"]
sp["B8"] = "=VAL_DATE"
d = "EDATE($B$8,12*$B$9)"
nxt = f"({d}+IF(WEEKDAY({d},2)=6,2,IF(WEEKDAY({d},2)=7,1,0)))"
prv = f"({d}-IF(WEEKDAY({d},2)=6,1,IF(WEEKDAY({d},2)=7,2,0)))"
sp["B10"] = f"=IF(MONTH({nxt})<>MONTH({d}),{prv},{nxt})"
wb.calculation.fullCalcOnLoad = True
wb.save(WB)
print("F22 -> settle basis; Swap_Pricer effective=settle, maturity BD-rolled")
