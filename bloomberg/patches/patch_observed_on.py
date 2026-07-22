"""
Use the OBSERVED overnight rate, not one fitted to the answer.

The o/n stub had been set to 3.64055%, obtained by minimising the error against
the S490 discount column. That is circular: an input tuned until our output
matched the mandate output, which stops the comparison being an independent
check.

The S490 screen shows an observed overnight rate of 3.59000 (bid = ask, i.e. a
fixing rather than a quote). Using it:

    observed 3.59%     max |zero diff| 0.881bp  (front end), 0.106bp from 2Y
    fitted   3.64055%  max |zero diff| 0.258bp  (front end), 0.087bp from 2Y

The fit bought 0.6bp at the very front and nothing beyond 2Y. Not worth the
circularity - and on a terminal SOFR_Fixings!C6 pulls the real fixing anyway, so
the observed value also keeps the off-terminal and live curves consistent.

INPUTS ARE NOW: the 32 OIS swap rates (bid/ask mid) + the overnight fixing.
Nothing is derived from the S490 zero/discount columns, which are purely the
comparison target.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font

WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg_Pricer.xlsx"
wb = load_workbook(WB)
f = wb["SOFR_Fixings"]
f["E6"].value = 3.59
f["E6"].number_format = "0.00000"
f["F6"].value = ("OBSERVED overnight rate from the S490 screen 07/21/26 (bid = ask = "
                 "3.59000), NOT fitted. Off-terminal fallback only: on a terminal C6 "
                 "pulls the live SOFRRATE fixing. Every curve input is now something we "
                 "pull - nothing is derived from the S490 zero/discount columns, which "
                 "are purely the comparison target.")
f["F6"].font = Font(name="Calibri", size=9, italic=True, color="666666")
wb.calculation.fullCalcOnLoad = True
wb.save(WB)
print("o/n stub set to the observed 3.59%")
