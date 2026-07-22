"""
Wire the interpolation switch into the copies that are ACTUALLY used.

Curve_Interface advertises itself as "one function used by both Swap_Pricer and
the CDS module", but nothing references its column G - a grep for references to
Curve_Interface!G returns zero. The interpolation is triplicated as inline
formulas:

    Curve_Interface!G        display only, dead
    CDS_Schedule!G, !H       DF(end), DF(mid) - what the CDS pricer actually uses
    Swap_Pricer!B19, !B20    DF(effective), DF(maturity)

So a switch on Curve_Interface alone changes nothing. This puts the same branch
into the live copies, all reading the one selector at Curve_Interface!J7.

Both branches interpolate off the same pillar grid (Curve_Interface K/L for the
CDS, Swap_Pricer K/L for the swap):

    Step-function forward  DF = DF_i * (DF_i+1/DF_i)^((d-K_i)/(K_i+1-K_i))
    Piecewise Linear (Simple)  DF = 1/(1+r_s*t), r_s linear on ACT/360, flat
                               before the first instrument maturity
"""

from openpyxl import load_workbook
from openpyxl.styles import Font

WB = "/Users/nigelli/Desktop/openusdcurve/bloomberg/USD_SOFR_Curve_Bloomberg_Pricer.xlsx"
MODE = "Curve_Interface!$J$7"
NOTE = Font(name="Calibri", size=9, italic=True, color="666666")


def branches(d, K, L):
    """Return (step_fn_forward, piecewise_linear_simple) for date expression d."""
    i = f"MATCH({d},{K},1)"
    Li, Li1 = f"INDEX({L},{i})", f"INDEX({L},{i}+1)"
    Ki, Ki1 = f"INDEX({K},{i})", f"INDEX({K},{i}+1)"
    m3 = f"{Li}*({Li1}/{Li})^(({d}-{Ki})/({Ki1}-{Ki}))"
    tp, tn, tt = f"(({Ki}-VAL_DATE)/360)", f"(({Ki1}-VAL_DATE)/360)", f"(({d}-VAL_DATE)/360)"
    rB = f"IF({tn}<=0,0,(1/{Li1}-1)/{tn})"
    rA = f"IF({tp}<=0,{rB},(1/{Li}-1)/{tp})"
    w = f"IF({tn}-{tp}=0,0,({tt}-{tp})/({tn}-{tp}))"
    m1 = f"IF({tt}<=0,1,1/(1+({rA}+({rB}-{rA})*{w})*{tt}))"
    return m3, m1


def wrap(d, K, L):
    m3, m1 = branches(d, K, L)
    fallback = f"INDEX({L},MATCH({d},{K},1))"
    return (f'=IFERROR(IF({MODE}="Piecewise Linear (Simple)",{m1},{m3}),{fallback})')


def main():
    wb = load_workbook(WB)

    # ---- CDS_Schedule: DF(end) on the pay date, DF(mid) on the period midpoint
    cs = wb["CDS_Schedule"]
    K = "'Curve_Interface'!$K$8:$K$73"
    L = "'Curve_Interface'!$L$8:$L$73"
    n = 0
    for r in range(7, 47):
        if cs[f"C{r}"].value is None:
            continue
        cs[f"G{r}"] = wrap(f"C{r}", K, L)
        cs[f"H{r}"] = wrap(f"(B{r}+C{r})/2", K, L)
        n += 1
    cs["P5"] = ("DF(end)/DF(mid) honour the interpolation selector at "
                "Curve_Interface!J7. Default is step-function forward, which is the "
                "method verified against the S490 screen.")
    cs["P5"].font = NOTE

    # ---- Swap_Pricer: its own pillar grid on its own sheet
    sp = wb["Swap_Pricer"]
    K2, L2 = "$K$6:$K$71", "$L$6:$L$71"
    sp["B19"] = wrap("$B$8", K2, L2)
    sp["B20"] = wrap("$B$10", K2, L2)

    wb.calculation.fullCalcOnLoad = True
    wb.save(WB)
    print(f"wired: CDS_Schedule {n} rows (G+H), Swap_Pricer B19/B20")


if __name__ == "__main__":
    main()
