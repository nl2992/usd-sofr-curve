Attribute VB_Name = "CDSStrip"
'==============================================================================
' CDSStrip - the hazard strip, self-contained. Replaces the Hazard_Solver sheet.
'
' Hazard_Solver was 3,712 cells: it sliced CDS_Schedule into per-tenor row
' vectors and ran a 30-step bisection ladder on each. This does both in one call.
'
'   CDS_StripHazard(k, method, mats, spreads, recovery, aiNet,
'                   payDates, alphas, dts, dfEnd, dfMid)
'
'   k        which tenor, 1..N
'   method   "" or "BRENT" (default), "BISECTION", "FALSE POSITION", "SECANT",
'            "NEWTON", "HALLEY", "HOUSEHOLDER", "RIDDERS"
'
' Ranges come straight off the sheets:
'   mats     Hazard_Bootstrap!$B$7:$B$12     spreads  Hazard_Bootstrap!$C$7:$C$12
'   recovery CDS_Parameters!$B$8             aiNet    CDS_Parameters!$B$26
'   payDates CDS_Schedule!$C$7:$C$46         alphas   CDS_Schedule!$D$7:$D$46
'   dts      CDS_Schedule!$E$7:$E$46         dfEnd    CDS_Schedule!$G$7:$G$46
'   dfMid    CDS_Schedule!$H$7:$H$46
'
' Safe against circularity: D, E, G and H are accruals and discount factors and
' none depend on the hazard column. Column I does - never pass it.
'
' Scheme is the white paper's, section 4 p.8: hazard piecewise constant, solved
' maturity by maturity, each segment a 1-D root-find with earlier hazards held
' fixed. The paper does not name a method, which is why this one is a parameter.
' Derivatives for Newton/Halley/Householder are analytic - dQ/dh = -dt*Q.
'
' The whole strip is computed once per input set and cached, so six cells cost
' one pass. CDS_StripIterations() reports the last segment's iteration count.
'==============================================================================
Option Explicit

Private Const XTOL As Double = 0.000000000000001
Private Const MAXIT As Long = 200

Private cKey As String
Private cH(1 To 32) As Double
Private mIter As Long

' segment state, set before each solve so the objective can see it
Private gS As Double, gR As Double, gAI As Double
Private gQ0 As Double, gPF As Double, gRF As Double
Private gLo As Long, gHi As Long
Private gA As Variant, gD As Variant, gE As Variant, gM As Variant


Public Function CDS_StripIterations() As Long
    CDS_StripIterations = mIter
End Function


Public Function CDS_StripHazard(ByVal k As Long, ByVal method As String, _
        ByVal mats As Range, ByVal spreads As Range, ByVal recovery As Double, _
        ByVal aiNet As Double, ByVal payDates As Range, ByVal alphas As Range, _
        ByVal dts As Range, ByVal dfEnd As Range, ByVal dfMid As Range) As Variant

    Dim nT As Long, nP As Long, i As Long, j As Long
    Dim key As String, h As Double, q As Double, qn As Double, d As Double
    Dim i0 As Long, iLo As Long, iHi As Long, mth As String

    nT = mats.Cells.Count: nP = alphas.Cells.Count
    If k < 1 Or k > nT Then CDS_StripHazard = CVErr(xlErrValue): Exit Function
    mth = UCase$(Trim$(method))
    If Len(mth) = 0 Then mth = "BRENT"

    gA = alphas.Value2: gD = dts.Value2: gE = dfEnd.Value2: gM = dfMid.Value2
    gR = recovery: gAI = aiNet

    key = mth & "|" & CStr(recovery) & "|" & CStr(aiNet)
    For i = 1 To nT
        key = key & "|" & CStr(NumOf(mats.Cells(i).Value)) & ":" & CStr(NumOf(spreads.Cells(i).Value))
    Next i
    For i = 1 To nP
        key = key & "," & Format$(Num2(gE, i), "0.000000000")
    Next i

    If key <> cKey Then
        gQ0 = 1#: gPF = 0#: gRF = 0#: i0 = 1
        For j = 1 To nT
            iLo = i0: iHi = i0 - 1
            For i = i0 To nP
                If Num2(gA, i) > 0# And _
                   NumOf(payDates.Cells(i).Value) <= NumOf(mats.Cells(j).Value) + 0.5 Then
                    iHi = i
                End If
            Next i
            gS = NumOf(spreads.Cells(j).Value) / 10000#
            gLo = iLo: gHi = iHi
            h = Solve(mth)
            cH(j) = h
            q = gQ0
            For i = iLo To iHi
                If Num2(gA, i) > 0# Then
                    qn = q * Exp(-h * Num2(gD, i)): d = q - qn
                    gPF = gPF + Num2(gM, i) * d
                    gRF = gRF + Num2(gA, i) * Num2(gE, i) * qn + 0.5 * Num2(gA, i) * Num2(gM, i) * d
                    q = qn
                End If
            Next i
            gQ0 = q
            If iHi >= iLo Then i0 = iHi + 1
        Next j
        cKey = key
    End If
    CDS_StripHazard = cH(k)
End Function


'--- objective and analytic derivatives for the current segment ----------------
Private Function F(ByVal h As Double) As Double
    Dim i As Long, q As Double, qn As Double, d As Double, a As Double
    Dim prot As Double, rpv As Double
    prot = gPF: rpv = gRF: q = gQ0
    For i = gLo To gHi
        a = Num2(gA, i)
        If a > 0# Then
            qn = q * Exp(-h * Num2(gD, i)): d = q - qn
            prot = prot + Num2(gM, i) * d
            rpv = rpv + a * Num2(gE, i) * qn + 0.5 * a * Num2(gM, i) * d
            q = qn
        End If
    Next i
    F = (1# - gR) * prot - gS * (rpv - gAI)
End Function

Private Function D1(ByVal h As Double) As Double
    Dim i As Long, q As Double, dq As Double, dqP As Double, a As Double, t As Double
    Dim dP As Double, dR As Double
    q = gQ0: dqP = 0#
    For i = gLo To gHi
        a = Num2(gA, i)
        If a > 0# Then
            t = Num2(gD, i)
            q = gQ0 * Exp(-h * CumTo(i))
            dq = -CumTo(i) * q
            dP = dP + Num2(gM, i) * (dqP - dq)
            dR = dR + a * Num2(gE, i) * dq + 0.5 * a * Num2(gM, i) * (dqP - dq)
            dqP = dq
        End If
    Next i
    D1 = (1# - gR) * dP - gS * dR
End Function

Private Function D2(ByVal h As Double) As Double
    Dim i As Long, q As Double, d2q As Double, d2qP As Double, a As Double, c As Double
    Dim dP As Double, dR As Double
    d2qP = 0#
    For i = gLo To gHi
        a = Num2(gA, i)
        If a > 0# Then
            c = CumTo(i)
            q = gQ0 * Exp(-h * c)
            d2q = c * c * q
            dP = dP + Num2(gM, i) * (d2qP - d2q)
            dR = dR + a * Num2(gE, i) * d2q + 0.5 * a * Num2(gM, i) * (d2qP - d2q)
            d2qP = d2q
        End If
    Next i
    D2 = (1# - gR) * dP - gS * dR
End Function

' cumulative dt from the segment start to period i
Private Function CumTo(ByVal i As Long) As Double
    Dim j As Long, s As Double
    For j = gLo To i
        If Num2(gA, j) > 0# Then s = s + Num2(gD, j)
    Next j
    CumTo = s
End Function


'--- solvers -------------------------------------------------------------------
Private Function Solve(ByVal mth As String) As Double
    Dim a As Double, b As Double, fa As Double, fb As Double
    mIter = 0
    a = 0#: b = 3#
    fa = F(a): fb = F(b)
    If fa * fb > 0# Then Solve = 0#: Exit Function      ' p.9 not strippable

    Select Case mth
        Case "BISECTION":      Solve = MBis(a, b)
        Case "FALSE POSITION": Solve = MFal(a, b)
        Case "SECANT":         Solve = MSec(a, b)
        Case "NEWTON":         Solve = MHou(1, (a + b) / 2#)
        Case "HALLEY":         Solve = MHou(2, (a + b) / 2#)
        Case "HOUSEHOLDER":    Solve = MHou(3, (a + b) / 2#)
        Case "RIDDERS":        Solve = MRid(a, b)
        Case Else:             Solve = MBre(a, b)
    End Select
End Function

Private Function MBis(ByVal a As Double, ByVal b As Double) As Double
    Dim m As Double, fa As Double, fm As Double
    fa = F(a)
    Do While mIter < MAXIT
        mIter = mIter + 1
        m = (a + b) / 2#: fm = F(m)
        If fm = 0# Or (b - a) / 2# < XTOL Then Exit Do
        If fa * fm < 0# Then b = m Else a = m: fa = fm
    Loop
    MBis = m
End Function

Private Function MFal(ByVal a As Double, ByVal b As Double) As Double
    Dim c As Double, cO As Double, fa As Double, fb As Double, fc As Double
    fa = F(a): fb = F(b): cO = a
    Do While mIter < MAXIT
        mIter = mIter + 1
        c = (a * fb - b * fa) / (fb - fa): fc = F(c)
        If fc = 0# Or Abs(c - cO) < XTOL Then Exit Do
        If fa * fc < 0# Then b = c: fb = fc Else a = c: fa = fc
        cO = c
    Loop
    MFal = c
End Function

Private Function MSec(ByVal x0 As Double, ByVal x1 As Double) As Double
    Dim f0 As Double, f1 As Double, x2 As Double
    f0 = F(x0): f1 = F(x1)
    Do While mIter < MAXIT
        mIter = mIter + 1
        If f1 = f0 Then Exit Do
        x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        If Abs(x2 - x1) < XTOL Then x1 = x2: Exit Do
        x0 = x1: f0 = f1: x1 = x2: f1 = F(x1)
    Loop
    MSec = x1
End Function

Private Function MHou(ByVal d As Long, ByVal x As Double) As Double
    Dim f0 As Double, f1 As Double, f2 As Double, st As Double
    Do While mIter < MAXIT
        mIter = mIter + 1
        f0 = F(x): f1 = D1(x)
        If f1 = 0# Then Exit Do
        If d <= 1 Then
            st = -f0 / f1
        Else
            f2 = D2(x)
            st = -(2# * f0 * f1) / (2# * f1 * f1 - f0 * f2)
        End If
        x = x + st
        If x < 0# Then x = 0.0000000001
        If Abs(st) < XTOL Then Exit Do
    Loop
    MHou = x
End Function

Private Function MRid(ByVal a As Double, ByVal b As Double) As Double
    Dim fa As Double, fb As Double, fm As Double, fn As Double
    Dim m As Double, xn As Double, s As Double, xo As Double
    fa = F(a): fb = F(b): xo = a
    Do While mIter < MAXIT
        mIter = mIter + 1
        m = (a + b) / 2#: fm = F(m)
        s = Sqr(fm * fm - fa * fb)
        If s = 0# Then MRid = m: Exit Function
        xn = m + (m - a) * Sgn(fa - fb) * fm / s
        fn = F(xn)
        If fn = 0# Or Abs(xn - xo) < XTOL Then MRid = xn: Exit Function
        xo = xn
        If fm * fn < 0# Then
            a = m: fa = fm: b = xn: fb = fn
        ElseIf fa * fn < 0# Then
            b = xn: fb = fn
        Else
            a = xn: fa = fn
        End If
    Loop
    MRid = xn
End Function

Private Function MBre(ByVal a As Double, ByVal b As Double) As Double
    Dim c As Double, d As Double, s As Double, t As Double
    Dim fa As Double, fb As Double, fc As Double, fs As Double
    Dim mflag As Boolean, hasD As Boolean
    fa = F(a): fb = F(b)
    If Abs(fa) < Abs(fb) Then
        t = a: a = b: b = t
        t = fa: fa = fb: fb = t
    End If
    c = a: fc = fa: mflag = True
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
        fs = F(s)
        d = c: hasD = True: c = b: fc = fb
        If fa * fs < 0# Then b = s: fb = fs Else a = s: fa = fs
        If Abs(fa) < Abs(fb) Then
            t = a: a = b: b = t
            t = fa: fa = fb: fb = t
        End If
    Loop
    MBre = b
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
    Acc = True
End Function

Private Function Num2(ByRef v As Variant, ByVal i As Long) As Double
    On Error Resume Next
    If IsNumeric(v(i, 1)) Then Num2 = CDbl(v(i, 1))
    On Error GoTo 0
End Function

Private Function NumOf(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = 0#
End Function
