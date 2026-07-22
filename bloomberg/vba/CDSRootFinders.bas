Attribute VB_Name = "CDSRootFinders"
'==============================================================================
' CDSRootFinders - the MATH5030 M2 methods applied to the CDS hazard strip
'
' The objective is the par-spread condition, B-Model (3.3):
'
'   f(h) = (1-R)*Prot(h) - S*( RPV01(h) - AI )
'
'   Prot(h)  = ProtFixed  + SUM  zMid * (Q(i-1) - Q(i))
'   RPV01(h) = RPV01Fixed + SUM  a*zEnd*Q(i) + 1/2 SUM a*zMid*(Q(i-1) - Q(i))
'   Q(i)     = Qstart * exp( -h * cum(i) )
'
' BASELINE. The white paper (section 4, p.8) fixes the SCHEME, not the algorithm:
' piecewise-constant hazard, solved maturity by maturity, each h_i+1 found by a
' one-dimensional root-find on (T_Mi, T_Mi+1] with all earlier hazards held
' fixed. It does not name a method. Every method below solves that same f(h) at
' that same step, so they are interchangeable and must agree - what differs is
' cost and robustness, not the answer. Our in-cell Hazard_Solver is bisection.
'
' Derivatives are analytic, not bumped. Q(i) = Qstart*exp(-h*cum(i)), so
'   dQ/dh = -cum*Q      d2Q/dh2 = cum^2*Q
' which makes Newton, Halley and Householder exact rather than approximate.
'
' Methods, matching M2:
'   bisection        bracketing, linear (halves each step)     - the safe floor
'   false position   bracketing, super-linear                  - can stall one-sided
'   secant           open, order ~1.618, no derivative
'   Newton           open, order 2, needs f'
'   Halley           open, order 3, needs f' and f''
'   Householder d    open, order d+1  (d=1 Newton, d=2 Halley)
'   Ridders          bracketing hybrid, exponential interpolation
'   Brent            bracketing hybrid, IQI + secant + bisection
'
' Stopping is the M2 mixed criterion:  |dx| < xtol + |x|*rtol.
'
' Import: Alt+F11 > File > Import File. Needs CDSBrent.bas alongside it for
' CDS_Objective. Save as .xlsm.
'==============================================================================
Option Explicit

Private Const XTOL As Double = 0.000000000000001
Private Const RTOL As Double = 0.000000000000001
Private Const MAXIT As Long = 200
Private mIter As Long

Public Function CDS_RootIterations() As Long
    CDS_RootIterations = mIter
End Function


'--- analytic first derivative of f -------------------------------------------
Public Function CDS_ObjectiveD1(ByVal h As Double, ByVal S As Double, ByVal R As Double, _
        ByVal qStart As Double, ByVal aiNet As Double, ByVal alphas As Range, _
        ByVal cumDts As Range, ByVal dfEnd As Range, ByVal dfMid As Range) As Double
    Dim i As Long, a As Double, cum As Double, ze As Double, zm As Double
    Dim q As Double, dq As Double, qP As Double, dqP As Double
    Dim dProt As Double, dRpv As Double, cumP As Double
    qP = qStart: dqP = 0#: cumP = 0#
    For i = 1 To alphas.Cells.Count
        a = NumOf(alphas.Cells(i).Value)
        If a <> 0 Then
            cum = NumOf(cumDts.Cells(i).Value)
            ze = NumOf(dfEnd.Cells(i).Value)
            zm = NumOf(dfMid.Cells(i).Value)
            q = qStart * Exp(-h * cum)
            dq = -cum * q
            dProt = dProt + zm * (dqP - dq)
            dRpv = dRpv + a * ze * dq + 0.5 * a * zm * (dqP - dq)
            qP = q: dqP = dq: cumP = cum
        End If
    Next i
    CDS_ObjectiveD1 = (1# - R) * dProt - S * dRpv
End Function


'--- analytic second derivative ------------------------------------------------
Public Function CDS_ObjectiveD2(ByVal h As Double, ByVal S As Double, ByVal R As Double, _
        ByVal qStart As Double, ByVal aiNet As Double, ByVal alphas As Range, _
        ByVal cumDts As Range, ByVal dfEnd As Range, ByVal dfMid As Range) As Double
    Dim i As Long, a As Double, cum As Double, ze As Double, zm As Double
    Dim q As Double, d2q As Double, d2qP As Double
    Dim d2Prot As Double, d2Rpv As Double
    d2qP = 0#
    For i = 1 To alphas.Cells.Count
        a = NumOf(alphas.Cells(i).Value)
        If a <> 0 Then
            cum = NumOf(cumDts.Cells(i).Value)
            ze = NumOf(dfEnd.Cells(i).Value)
            zm = NumOf(dfMid.Cells(i).Value)
            q = qStart * Exp(-h * cum)
            d2q = cum * cum * q
            d2Prot = d2Prot + zm * (d2qP - d2q)
            d2Rpv = d2Rpv + a * ze * d2q + 0.5 * a * zm * (d2qP - d2q)
            d2qP = d2q
        End If
    Next i
    CDS_ObjectiveD2 = (1# - R) * d2Prot - S * d2Rpv
End Function


'==============================================================================
' CDS_Root - one entry point, pick the method by name
'   method: "BISECTION" | "FALSE POSITION" | "SECANT" | "NEWTON" | "HALLEY"
'           | "HOUSEHOLDER" | "RIDDERS" | "BRENT"
' Read the iteration count with CDS_RootIterations() straight after.
'==============================================================================
Public Function CDS_Root(ByVal method As String, ByVal S As Double, ByVal R As Double, _
        ByVal qStart As Double, ByVal protFixed As Double, ByVal rpv01Fixed As Double, _
        ByVal aiNet As Double, ByVal alphas As Range, ByVal cumDts As Range, _
        ByVal dfEnd As Range, ByVal dfMid As Range, _
        Optional ByVal hOrder As Long = 3) As Variant

    Dim a As Double, b As Double, fa As Double, fb As Double
    mIter = 0
    a = 0#: b = 3#
    fa = F(a, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
    fb = F(b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)

    ' p.9 strippability: no sign change means the quote is outside the achievable
    ' interval for this segment. Report it, do not return a bracket end.
    If fa * fb > 0# Then CDS_Root = CVErr(xlErrNum): Exit Function

    Select Case UCase$(Trim$(method))
        Case "BISECTION":      CDS_Root = MBisect(a, b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "FALSE POSITION": CDS_Root = MFalsePos(a, b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "SECANT":         CDS_Root = MSecant(a, b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "NEWTON":         CDS_Root = MHouse(1, (a + b) / 2#, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "HALLEY":         CDS_Root = MHouse(2, (a + b) / 2#, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "HOUSEHOLDER":    CDS_Root = MHouse(hOrder, (a + b) / 2#, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "RIDDERS":        CDS_Root = MRidders(a, b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case "BRENT":          CDS_Root = MBrent(a, b, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
        Case Else:             CDS_Root = CVErr(xlErrValue)
    End Select
End Function


'==============================================================================
' methods
'==============================================================================
Private Function MBisect(ByVal a As Double, ByVal b As Double, ParamArray p() As Variant) As Double
    Dim m As Double, fa As Double, fm As Double
    fa = FA(a, p)
    Do While mIter < MAXIT
        mIter = mIter + 1
        m = (a + b) / 2#
        fm = FA(m, p)
        If fm = 0# Or Conv((b - a) / 2#, m) Then Exit Do
        If fa * fm < 0# Then b = m Else a = m: fa = fm
    Loop
    MBisect = m
End Function

Private Function MFalsePos(ByVal a As Double, ByVal b As Double, ParamArray p() As Variant) As Double
    Dim c As Double, cOld As Double, fa As Double, fb As Double, fc As Double
    fa = FA(a, p): fb = FA(b, p): cOld = a
    Do While mIter < MAXIT
        mIter = mIter + 1
        c = (a * fb - b * fa) / (fb - fa)
        fc = FA(c, p)
        If fc = 0# Or Conv(c - cOld, c) Then Exit Do
        If fa * fc < 0# Then
            b = c: fb = fc
        Else
            a = c: fa = fc
        End If
        cOld = c
    Loop
    MFalsePos = c
End Function

Private Function MSecant(ByVal x0 As Double, ByVal x1 As Double, ParamArray p() As Variant) As Double
    Dim f0 As Double, f1 As Double, x2 As Double
    f0 = FA(x0, p): f1 = FA(x1, p)
    Do While mIter < MAXIT
        mIter = mIter + 1
        If f1 = f0 Then Exit Do
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        If Conv(x2 - x1, x2) Then x1 = x2: Exit Do
        x0 = x1: f0 = f1: x1 = x2: f1 = FA(x1, p)
    Loop
    MSecant = x1
End Function

' Householder of order d.  d=1 is Newton, d=2 is Halley - the M2 identity.
Private Function MHouse(ByVal d As Long, ByVal x As Double, ParamArray p() As Variant) As Double
    Dim f0 As Double, f1 As Double, f2 As Double, step As Double
    Do While mIter < MAXIT
        mIter = mIter + 1
        f0 = FA(x, p)
        f1 = D1A(x, p)
        If f1 = 0# Then Exit Do
        If d <= 1 Then
            step = -f0 / f1                                   ' Newton
        Else
            f2 = D2A(x, p)
            step = -(2# * f0 * f1) / (2# * f1 * f1 - f0 * f2) ' Halley / d>=2
        End If
        x = x + step
        If x < 0# Then x = 0.0000000001
        If Conv(step, x) Then Exit Do
    Loop
    MHouse = x
End Function

Private Function MRidders(ByVal a As Double, ByVal b As Double, ParamArray p() As Variant) As Double
    Dim fa As Double, fb As Double, fm As Double, fnew As Double
    Dim m As Double, xnew As Double, s As Double, xold As Double
    fa = FA(a, p): fb = FA(b, p): xold = a
    Do While mIter < MAXIT
        mIter = mIter + 1
        m = (a + b) / 2#
        fm = FA(m, p)
        s = Sqr(fm * fm - fa * fb)
        If s = 0# Then MRidders = m: Exit Function
        xnew = m + (m - a) * Sgn(fa - fb) * fm / s
        fnew = FA(xnew, p)
        If fnew = 0# Or Conv(xnew - xold, xnew) Then MRidders = xnew: Exit Function
        xold = xnew
        If fm * fnew < 0# Then
            a = m: fa = fm: b = xnew: fb = fnew
        ElseIf fa * fnew < 0# Then
            b = xnew: fb = fnew
        Else
            a = xnew: fa = fnew
        End If
    Loop
    MRidders = xnew
End Function

Private Function MBrent(ByVal a As Double, ByVal b As Double, ParamArray p() As Variant) As Double
    Dim c As Double, d As Double, s As Double
    Dim fa As Double, fb As Double, fc As Double, fs As Double
    Dim mflag As Boolean, hasD As Boolean, t As Double
    fa = FA(a, p): fb = FA(b, p)
    If Abs(fa) < Abs(fb) Then
        t = a: a = b: b = t
        t = fa: fa = fb: fb = t
    End If
    c = a: fc = fa: mflag = True: hasD = False
    Do While fb <> 0# And Abs(b - a) > XTOL And mIter < MAXIT
        mIter = mIter + 1
        If (fa <> fc) And (fb <> fc) Then
            s = a * fb * fc / ((fa - fb) * (fa - fc)) _
              + b * fa * fc / ((fb - fa) * (fb - fc)) _
              + c * fa * fb / ((fc - fa) * (fc - fb))
        Else
            s = b - fb * (b - a) / (fb - fa)
        End If
        If Not Acc(s, a, b, c, d, hasD, mflag) Then
            s = (a + b) / 2#: mflag = True
        Else
            mflag = False
        End If
        fs = FA(s, p)
        d = c: hasD = True
        c = b: fc = fb
        If fa * fs < 0# Then b = s: fb = fs Else a = s: fa = fs
        If Abs(fa) < Abs(fb) Then
            t = a: a = b: b = t
            t = fa: fa = fb: fb = t
        End If
    Loop
    MBrent = b
End Function


'==============================================================================
' plumbing
'==============================================================================
Private Function Conv(ByVal dx As Double, ByVal x As Double) As Boolean
    Conv = (Abs(dx) < XTOL + Abs(x) * RTOL)      ' M2 mixed criterion
End Function

Private Function F(ByVal h As Double, ByVal S As Double, ByVal R As Double, _
        ByVal qStart As Double, ByVal protFixed As Double, ByVal rpv01Fixed As Double, _
        ByVal aiNet As Double, ByVal alphas As Range, ByVal cumDts As Range, _
        ByVal dfEnd As Range, ByVal dfMid As Range) As Double
    F = CDS_Objective(h, S, R, qStart, protFixed, rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)
End Function

Private Function FA(ByVal h As Double, ByRef p As Variant) As Double
    FA = CDS_Objective(h, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)), CDbl(p(3)), CDbl(p(4)), _
                       CDbl(p(5)), p(6), p(7), p(8), p(9))
End Function

Private Function D1A(ByVal h As Double, ByRef p As Variant) As Double
    D1A = CDS_ObjectiveD1(h, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)), CDbl(p(5)), _
                          p(6), p(7), p(8), p(9))
End Function

Private Function D2A(ByVal h As Double, ByRef p As Variant) As Double
    D2A = CDS_ObjectiveD2(h, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)), CDbl(p(5)), _
                          p(6), p(7), p(8), p(9))
End Function

Private Function Acc(ByVal s As Double, ByVal a As Double, ByVal b As Double, _
                     ByVal c As Double, ByVal d As Double, ByVal hasD As Boolean, _
                     ByVal mflag As Boolean) As Boolean
    Dim lo As Double, hi As Double, t As Double
    lo = (3# * a + b) / 4#: hi = b
    If lo > hi Then t = lo: lo = hi: hi = t
    If s <= lo Or s >= hi Then Acc = False: Exit Function
    If mflag And Abs(s - b) >= Abs(b - c) / 2# Then Acc = False: Exit Function
    If (Not mflag) And hasD And Abs(s - b) >= Abs(c - d) / 2# Then Acc = False: Exit Function
    If mflag And Abs(b - c) < XTOL Then Acc = False: Exit Function
    If (Not mflag) And hasD And Abs(c - d) < XTOL Then Acc = False: Exit Function
    Acc = True
End Function

Private Function NumOf(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = 0#
End Function
