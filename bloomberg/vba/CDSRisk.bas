Attribute VB_Name = "CDSRisk"
'==============================================================================
' CDSRisk - risk measures by full revalue, the way CDSW computes them.
'
' The analytic first-order numbers we had are a different quantity from
' Bloomberg's, which is why Rec01 came out -4,792 against their +72.64. White
' paper p.9: Spread DV01 and IR DV01 "represent the change in the value of the
' transaction as a result of a parallel shift of 1bp", holding all other inputs
' constant, and the recovery figure is the sensitivity to a 1% shift. That means
' bump, RE-STRIP the hazard curve so it still matches the quotes, and reprice.
' Holding the hazards fixed - which is what an analytic RPV01*1bp does - answers
' a different question.
'
'   CDS_MarketValue(bumpSprdBp, bumpDfBp, bumpRec,
'                   mats, spreads, recovery, aiNet,
'                   payDates, alphas, dts, tYears, dfEnd, dfMid,
'                   notional, couponBp, matDate)
'
' Returns the market value to a PROTECTION BUYER: protection leg minus premium
' leg, both discounted to the pricing date, B-Model (3.1).
'
' Each risk number is then a difference of two calls, so the working stays on the
' sheet rather than inside the function:
'
'   Spread DV01 = MV(+1bp on every quote) - MV(base)
'   IR DV01     = MV(+1bp on the discount curve) - MV(base)
'   Rec01       = MV(recovery + 1%) - MV(base)
'
' The discount bump is applied as D -> D * exp(-0.0001 * t), which is a parallel
' 1bp shift in the continuously-compounded zero curve. tYears is CDS_Schedule!F;
' the midpoint uses t - dt/2.
'
' Every range is read through Value2, so Dates arrive as serial Doubles. Reading
' them with .Value and testing IsNumeric is what broke the strip - IsNumeric
' returns False for a Date.
'==============================================================================
Option Explicit

' Zero-argument probe, so a blank risk cell can be told apart from a missing
' module. CDS_MarketValue is wrapped in IFERROR on the sheet, which hides the
' difference between "module not loaded" and "the call errored".
Public Function CDS_RiskLoaded() As String
    CDS_RiskLoaded = "CDSRisk loaded"
End Function

Public Function CDS_MarketValue(ByVal bumpSprdBp As Double, ByVal bumpDfBp As Double, _
        ByVal bumpRec As Double, ByVal mats As Range, ByVal spreads As Range, _
        ByVal recovery As Double, ByVal aiNet As Double, ByVal payDates As Range, _
        ByVal alphas As Range, ByVal dts As Range, ByVal tYears As Range, _
        ByVal dfEnd As Range, ByVal dfMid As Range, ByVal notional As Double, _
        ByVal couponBp As Double, ByVal matDate As Double) As Variant

    Dim vMt, vSp, vPd, vA, vD, vT, vE, vM
    Dim nT As Long, nP As Long, i As Long, j As Long
    Dim R As Double, h() As Double, seg() As Long
    Dim dE() As Double, dM() As Double
    Dim q As Double, qn As Double, d As Double
    Dim prot As Double, prem As Double, rpv As Double
    Dim qPrev As Double, pF As Double, rF As Double, i0 As Long, iLo As Long, iHi As Long

    On Error GoTo Fail
    vMt = mats.Value2: vSp = spreads.Value2: vPd = payDates.Value2
    vA = alphas.Value2: vD = dts.Value2: vT = tYears.Value2
    vE = dfEnd.Value2: vM = dfMid.Value2
    nT = mats.Cells.Count: nP = alphas.Cells.Count
    R = recovery + bumpRec
    If R >= 1# Then R = 0.999999

    ' bumped discount factors: parallel 1bp on the continuous zero curve
    ReDim dE(1 To nP): ReDim dM(1 To nP)
    For i = 1 To nP
        dE(i) = N2(vE, i) * Exp(-bumpDfBp / 10000# * N2(vT, i))
        dM(i) = N2(vM, i) * Exp(-bumpDfBp / 10000# * (N2(vT, i) - N2(vD, i) / 2#))
    Next i

    ' re-strip against the bumped inputs
    ReDim h(1 To nT): ReDim seg(1 To nT)
    qPrev = 1#: pF = 0#: rF = 0#: i0 = 1
    For j = 1 To nT
        iLo = i0: iHi = i0 - 1
        For i = i0 To nP
            If N2(vA, i) > 0# And N2(vPd, i) <= N2(vMt, j) + 0.5 Then iHi = i
        Next i
        If iHi < iLo Then CDS_MarketValue = CVErr(xlErrNum): Exit Function
        h(j) = SolveSeg((N2(vSp, j) + bumpSprdBp) / 10000#, R, aiNet, qPrev, pF, rF, _
                        iLo, iHi, vA, vD, dE, dM)
        seg(j) = iHi
        q = qPrev
        For i = iLo To iHi
            qn = q * Exp(-h(j) * N2(vD, i)): d = q - qn
            pF = pF + dM(i) * d
            rF = rF + N2(vA, i) * dE(i) * qn + 0.5 * N2(vA, i) * dM(i) * d
            q = qn
        Next i
        qPrev = q: i0 = iHi + 1
    Next j

    ' price to maturity off the re-stripped curve
    q = 1#: j = 1
    For i = 1 To nP
        If N2(vPd, i) > matDate + 0.5 Then Exit For
        Do While j < nT And i > seg(j)
            j = j + 1
        Loop
        qn = q * Exp(-h(j) * N2(vD, i)): d = q - qn
        prot = prot + dM(i) * d
        rpv = rpv + N2(vA, i) * dE(i) * qn + 0.5 * N2(vA, i) * dM(i) * d
        q = qn
    Next i
    prem = couponBp / 10000# * rpv
    CDS_MarketValue = notional * ((1# - R) * prot - prem)
    Exit Function
Fail:
    CDS_MarketValue = CVErr(xlErrValue)
End Function


Private Function SolveSeg(ByVal S As Double, ByVal R As Double, ByVal ai As Double, _
        ByVal q0 As Double, ByVal pF As Double, ByVal rF As Double, _
        ByVal iLo As Long, ByVal iHi As Long, ByRef vA, ByRef vD, _
        ByRef dE() As Double, ByRef dM() As Double) As Double
    Dim a As Double, b As Double, m As Double, fa As Double, fm As Double, k As Long
    a = 0#: b = 3#
    fa = SegF(a, S, R, ai, q0, pF, rF, iLo, iHi, vA, vD, dE, dM)
    If fa * SegF(b, S, R, ai, q0, pF, rF, iLo, iHi, vA, vD, dE, dM) > 0# Then
        SolveSeg = 0#: Exit Function
    End If
    For k = 1 To 200
        m = (a + b) / 2#
        fm = SegF(m, S, R, ai, q0, pF, rF, iLo, iHi, vA, vD, dE, dM)
        If fm = 0# Or (b - a) / 2# < 0.000000000000001 Then Exit For
        If fa * fm < 0# Then b = m Else a = m: fa = fm
    Next k
    SolveSeg = m
End Function

Private Function SegF(ByVal hz As Double, ByVal S As Double, ByVal R As Double, _
        ByVal ai As Double, ByVal q0 As Double, ByVal pF As Double, ByVal rF As Double, _
        ByVal iLo As Long, ByVal iHi As Long, ByRef vA, ByRef vD, _
        ByRef dE() As Double, ByRef dM() As Double) As Double
    Dim i As Long, q As Double, qn As Double, d As Double, a As Double
    Dim prot As Double, rpv As Double
    prot = pF: rpv = rF: q = q0
    For i = iLo To iHi
        a = N2(vA, i)
        If a > 0# Then
            qn = q * Exp(-hz * N2(vD, i)): d = q - qn
            prot = prot + dM(i) * d
            rpv = rpv + a * dE(i) * qn + 0.5 * a * dM(i) * d
            q = qn
        End If
    Next i
    SegF = (1# - R) * prot - S * (rpv - ai)
End Function

Private Function N2(ByRef v As Variant, ByVal i As Long) As Double
    On Error Resume Next
    If IsNumeric(v(i, 1)) Then N2 = CDbl(v(i, 1))
    On Error GoTo 0
End Function
