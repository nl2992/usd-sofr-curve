"""
Fix the spot node of the curve grids.

Curve_Interface!L8 and Swap_Pricer!L6 hardcoded DF(spot) = 1, but every other
node on those grids comes from Bootstrap!H, which is discounted to VAL_DATE -
where DF(spot) = Bootstrap!D4 = 1/(1+r_on*(spot-VAL_DATE)/360), not 1.

The grid was therefore internally inconsistent by the 2-day spot stub. Effect on
a 1Y par swap: sheet 4.05108% vs correct 4.03032%, against a 1Y OIS input of
4.03030% - a 2.08bp error, worst at the short end and decaying with maturity.
"""
from openpyxl import load_workbook
WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg_Pricer.xlsx"
wb = load_workbook(WB)
for sheet, cell in (("Curve_Interface", "L8"), ("Swap_Pricer", "L6")):
    wb[sheet][cell] = "=Bootstrap!$D$4"
    wb[sheet][cell].number_format = "0.00000000"
wb.calculation.fullCalcOnLoad = True
wb.save(WB)
print("spot node set to Bootstrap!D4 on Curve_Interface!L8 and Swap_Pricer!L6")
