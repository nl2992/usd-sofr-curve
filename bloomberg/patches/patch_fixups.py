"""
Two corrections applied after the live-validation wiring.

1. Bloomberg_S490_Validation!H8 held a SECOND BDS(CURVE_TENOR_RATES) call, left
   over from the original sheet. It spills right and down, straight through the
   J12 dump block added for the live lookup - on a terminal the two arrays would
   have collided and corrupted each other. Removed.

2. SOFR_Fixings!F6 note rewritten to state what actually happens live: C6 pulls
   the real SOFRRATE fixing and OVERRIDES the fitted 3.64055% fallback, shifting
   the front end (~1.1bp at 1W, ~0.3bp at 1M, <0.03bp from 1Y out).
"""
from openpyxl import load_workbook
from openpyxl.styles import Font

WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg.xlsx"
wb = load_workbook(WB)

v = wb["Bloomberg_S490_Validation"]
v["H8"].value = ("(the CURVE_TENOR_RATES dump lives at J12 - a second copy here "
                 "would spill into it)")
v["H8"].font = Font(name="Calibri", size=9, italic=True, color="666666")

f = wb["SOFR_Fixings"]
f["F6"].value = ("LIVE on a terminal: C6 pulls the real SOFRRATE fixing and OVERRIDES "
                 "this cell. This 3.64055% is only the off-terminal fallback (fitted so "
                 "DFspot reproduces the 07/21/26 S490 screen). Expect the front end to "
                 "shift when live: ~1.1bp at 1W and ~0.3bp at 1M if the fixing is 3.59%, "
                 "but under 0.03bp from 1Y out. The long end is unaffected.")
f["F6"].font = Font(name="Calibri", size=9, italic=True, color="C00000")

wb.calculation.fullCalcOnLoad = True
wb.save(WB)
print("fixups applied")
