Attribute VB_Name = "CurveVBA"
'==============================================================================
' CurveVBA - the SOFR bootstrap as a VBA function
'
' Second, independent implementation of what the Bootstrap sheet does in cells.
' Test 2 on the Testing sheet runs THIS against the same quotes Test 1 feeds to
' the in-cell grid. Two implementations agreeing to ~1e-12 is a real check on
' both; a third copy of the same formulas would not have been.
'
' Deliberately transliterated, not improved. Every convention below exists
' because the cell grid does it that way, and matching Bloomberg is the point:
'
'   spot          = curve date + 2 business days   (Bootstrap!B4)
'   DF(spot)      = 1                              (Bootstrap!D4) - settle is
'                   the curve date on S490, so there is no spot-lag stub
'   <= 1Y         money-market:  DF = DFspot / (1 + S*tau0),  tau0 ACT/360
'                 measured from the CURVE date, not spot
'   > 1Y          par:  DF = (DFspot - S*A_prior) / (1 + S*tau_c),  tau_c ACT/360
'                 measured from spot
'   18M           priced as a pillar but NOT on the annual coupon schedule, so
'                 it adds nothing to the annuity and 2Y accrues 1Y->2Y.
'                 Bootstrap!I23 = 0 is where that lives. Getting this wrong
'                 double-counts a coupon and shifts the whole long end.
'   gap years     par rate linear in YEAR number between bracketing quotes;
'                 DF log-linear between bracketing pillar DFs - that is
'                 step-function forward, confirmed on S490 as "Step Forward (Cont)"
'   quoted pillar with gap years behind it: DF solved so the par identity holds
'                 WITH those interpolated gaps inside the annuity. Same job as
'                 the Curve_Solver sheet, by Brent instead of bisection - the
'                 residual is monotone so both find the same root.
'   dates         modified following, weekends only. NOT WORKDAY (Analysis
'                 ToolPak, evaluates unreliably here).
'   zero          continuous, ACT/365:  z = -ln(DF)/T * 100
'
' Usage:
'   =SOFR_Curve($D$5, $A$8:$A$39, $B$8:$B$39, A8, "ZERO")
'      curveDate, tenor labels, par rates in %, the tenor wanted, what to return
'   output: "ZERO" | "DF" | "DATE" | "PAR" | "T"
'
' The whole curve is rebuilt per call, then cached on the inputs, so 32 rows
' cost one bootstrap rather than 32.
'
' Import: Alt+F11 > File > Import File > this .bas. The workbook must be saved
' as .xlsm or the module will not persist. If Testing is moved into another
' workbook, this module has to be imported there too, or every VBA column
' returns #NAME?.
'==============================================================================
Option Explicit

Private Const MAXP As Long = 200
Private Const TOL As Double = 0.000000000000001   ' 1e-15

' cache
Private cKey As String
Private cN As Long
Private cTen() As String
Private cDate() As Double
Private cT() As Double
Private cPar() As Double
Private cDF() As Double


'--- the public entry point ----------------------------------------------------
Public Function SOFR_Curve(ByVal curveDate As Date, ByVal tenors As Range, _
                           ByVal rates As Range, ByVal wantTenor As String, _
                           ByVal output As String) As Variant
    Dim i As Long
    On Error GoTo Fail

    If Not BuildCurve(curveDate, tenors, rates) Then
        SOFR_Curve = CVErr(xlErrValue): Exit Function
    End If

    ' match on YEARS, not the label, so a 12M quote finds the 1Y pillar
    Dim wy As Double
    wy = TenorYears(wantTenor)
    If wy < 0 Then SOFR_Curve = CVErr(xlErrValue): Exit Function
    For i = 1 To cN
        If Abs(TenorYears(cTen(i)) - wy) < 0.0000001 Then
            Select Case UCase$(Trim$(output))
                Case "ZERO": SOFR_Curve = -Log(cDF(i)) / cT(i) * 100#
                Case "DF":   SOFR_Curve = cDF(i)
                Case "DATE": SOFR_Curve = CDate(cDate(i))
                Case "PAR":  SOFR_Curve = cPar(i)
                Case "T":    SOFR_Curve = cT(i)
                Case Else:   SOFR_Curve = CVErr(xlErrValue)
            End Select
            Exit Function
        End If
    Next i
    SOFR_Curve = CVErr(xlErrNA)     ' tenor not on the grid
    Exit Function
Fail:
    SOFR_Curve = CVErr(xlErrValue)
End Function


'--- build the whole curve, cached ---------------------------------------------
Private Function BuildCurve(ByVal curveDate As Date, ByVal tenors As Range, _
                            ByVal rates As Range) As Boolean
    Dim qTen() As String, qRate() As Double, qYr() As Double
    Dim nq As Long, i As Long, j As Long, k As Long
    Dim key As String, v As Variant, t As String
    Dim spot As Date, maxYr As Long, yr As Double
    Dim onSched() As Boolean, isQuoted() As Boolean
    Dim tau0() As Double, tauC() As Double, ann() As Double
    Dim annCum As Double, prevSched As Double, S As Double
    Dim lastQ As Long, nGap As Long

    BuildCurve = False
    If tenors.Cells.Count <> rates.Cells.Count Then Exit Function

    ' ---- collect the quotes, skipping blanks ----
    ReDim qTen(1 To tenors.Cells.Count)
    ReDim qRate(1 To tenors.Cells.Count)
    ReDim qYr(1 To tenors.Cells.Count)
    key = CStr(CLng(curveDate))
    nq = 0
    For i = 1 To tenors.Cells.Count
        t = UCase$(Trim$(CStr(tenors.Cells(i).Value)))
        v = rates.Cells(i).Value
        If Len(t) >= 2 And IsNumeric(v) Then
            nq = nq + 1
            qTen(nq) = t
            qRate(nq) = CDbl(v)
            qYr(nq) = TenorYears(t)
            If qYr(nq) < 0 Then Exit Function
            key = key & "|" & t & "=" & Format$(qRate(nq), "0.00000000")
        End If
    Next i
    If nq = 0 Then Exit Function

    If key = cKey Then BuildCurve = True: Exit Function   ' cache hit

    ' ---- build the grid: every quote, plus every integer year to the longest ----
    spot = AddBusDays(curveDate, 2)
    maxYr = 0
    For i = 1 To nq
        If qYr(i) > maxYr Then maxYr = Int(qYr(i) + 0.0000001)
    Next i

    ReDim cTen(1 To MAXP): ReDim cDate(1 To MAXP): ReDim cT(1 To MAXP)
    ReDim cPar(1 To MAXP): ReDim cDF(1 To MAXP)
    ReDim onSched(1 To MAXP): ReDim isQuoted(1 To MAXP)
    ReDim tau0(1 To MAXP): ReDim tauC(1 To MAXP): ReDim ann(1 To MAXP)
    cN = 0

    ' sub-year quotes first, in the order given
    For i = 1 To nq
        If qYr(i) < 1# Then
            cN = cN + 1: cTen(cN) = qTen(i): cPar(cN) = qRate(i)
            isQuoted(cN) = True: onSched(cN) = False
        End If
    Next i
    ' then each integer year, and after it any off-schedule quote (18M) that
    ' falls before the next year. Order matters: the annuity is walked in it.
    For yr = 1 To maxYr
        k = 0
        For i = 1 To nq
            If qYr(i) = yr Then k = i
        Next i
        cN = cN + 1
        cTen(cN) = CStr(yr) & "Y"
        onSched(cN) = True
        If k > 0 Then
            cPar(cN) = qRate(k): isQuoted(cN) = True
        Else
            cPar(cN) = InterpYear(yr, qYr, qRate, nq): isQuoted(cN) = False
        End If
        For i = 1 To nq
            If qYr(i) > yr And qYr(i) < yr + 1 Then
                cN = cN + 1: cTen(cN) = qTen(i): cPar(cN) = qRate(i)
                isQuoted(cN) = True
                onSched(cN) = False        ' Bootstrap!I23 = 0 - 18M is a pillar
            End If                          ' but not a coupon date
        Next i
    Next yr

    ' ---- dates, year fractions ----
    For i = 1 To cN
        cDate(i) = CDbl(TenorDate(curveDate, cTen(i)))
        cT(i) = (cDate(i) - CDbl(curveDate)) / 365#
        If TenorYears(cTen(i)) <= 1# Then
            tau0(i) = (cDate(i) - CDbl(curveDate)) / 360#      ' MM: from curve date
        Else
            tau0(i) = (cDate(i) - CDbl(spot)) / 360#           ' swaps: from spot
        End If
    Next i

    ' coupon accruals run between SCHEDULED pillars only
    prevSched = CDbl(curveDate)
    For i = 1 To cN
        If onSched(i) Then
            tauC(i) = (cDate(i) - prevSched) / 360#
            prevSched = cDate(i)
        Else
            ' off-schedule (18M): accrues from the previous scheduled pillar but
            ' never becomes one, so the next pillar still accrues from 1Y
            tauC(i) = (cDate(i) - prevSched) / 360#
        End If
    Next i

    ' ---- walk the pillars ----
    annCum = 0#
    lastQ = 0
    For i = 1 To cN
        S = cPar(i) / 100#
        If TenorYears(cTen(i)) <= 1# Then
            cDF(i) = 1# / (1# + S * tau0(i))          ' money market, DFspot = 1
        ElseIf Not isQuoted(i) Then
            ' gap year - DF is set by the solve at the next quoted pillar
        Else
            nGap = i - lastQ - 1
            If nGap <= 0 Then
                cDF(i) = (1# - S * annCum) / (1# + S * tauC(i))
            Else
                cDF(i) = SolvePillar(i, lastQ, S, annCum, tauC, onSched)
                FillGaps i, lastQ, tauC, onSched, annCum
            End If
        End If

        If isQuoted(i) Then
            If onSched(i) Then annCum = annCum + tauC(i) * cDF(i)
            lastQ = i
        End If
    Next i

    cKey = key
    BuildCurve = True
End Function


'--- solve a quoted pillar so the par identity holds WITH its gap rows ----------
Private Function SolvePillar(ByVal iB As Long, ByVal iA As Long, ByVal S As Double, _
                             ByVal annAnchor As Double, ByRef tauC() As Double, _
                             ByRef onSched() As Boolean) As Double
    Dim a As Double, b As Double, c As Double, d As Double, s2 As Double
    Dim fa As Double, fb As Double, fc As Double, fs As Double
    Dim mflag As Boolean, hasD As Boolean, it As Long, tmp As Double

    a = 0.000000001: b = 1#
    fa = Resid(a, iB, iA, S, annAnchor, tauC, onSched)
    fb = Resid(b, iB, iA, S, annAnchor, tauC, onSched)
    If fa * fb > 0# Then SolvePillar = 0#: Exit Function

    If Abs(fa) < Abs(fb) Then
        tmp = a: a = b: b = tmp
        tmp = fa: fa = fb: fb = tmp
    End If
    c = a: fc = fa: mflag = True: hasD = False: d = 0#

    Do While fb <> 0# And Abs(b - a) > TOL
        it = it + 1
        If it > 200 Then Exit Do
        If (fa <> fc) And (fb <> fc) Then
            s2 = a * fb * fc / ((fa - fb) * (fa - fc)) _
               + b * fa * fc / ((fb - fa) * (fb - fc)) _
               + c * fa * fb / ((fc - fa) * (fc - fb))
        Else
            s2 = b - fb * (b - a) / (fb - fa)
        End If
        If Not Accept(s2, a, b, c, d, hasD, mflag) Then
            s2 = (a + b) / 2#: mflag = True
        Else
            mflag = False
        End If
        fs = Resid(s2, iB, iA, S, annAnchor, tauC, onSched)
        d = c: hasD = True
        c = b: fc = fb
        If fa * fs < 0# Then
            b = s2: fb = fs
        Else
            a = s2: fa = fs
        End If
        If Abs(fa) < Abs(fb) Then
            tmp = a: a = b: b = tmp
            tmp = fa: fa = fb: fb = tmp
        End If
    Loop
    SolvePillar = b
End Function


' residual = DFspot - S*(annuity incl. interpolated gaps) - DF_B
Private Function Resid(ByVal dfB As Double, ByVal iB As Long, ByVal iA As Long, _
                       ByVal S As Double, ByVal annAnchor As Double, _
                       ByRef tauC() As Double, ByRef onSched() As Boolean) As Double
    Dim g As Long, tot As Double, dfg As Double
    tot = annAnchor
    For g = iA + 1 To iB - 1
        dfg = LogLin(dfB, iB, iA, g)
        If onSched(g) Then tot = tot + tauC(g) * dfg
    Next g
    tot = tot + tauC(iB) * dfB
    Resid = (1# - S * tot) - dfB
End Function


' log-linear in DF between the anchor and the trial pillar = step-forward
Private Function LogLin(ByVal dfB As Double, ByVal iB As Long, ByVal iA As Long, _
                        ByVal g As Long) As Double
    LogLin = cDF(iA) * (dfB / cDF(iA)) ^ ((cT(g) - cT(iA)) / (cT(iB) - cT(iA)))
End Function


Private Sub FillGaps(ByVal iB As Long, ByVal iA As Long, ByRef tauC() As Double, _
                     ByRef onSched() As Boolean, ByRef annCum As Double)
    Dim g As Long
    For g = iA + 1 To iB - 1
        cDF(g) = LogLin(cDF(iB), iB, iA, g)
        If onSched(g) Then annCum = annCum + tauC(g) * cDF(g)
    Next g
End Sub


Private Function Accept(ByVal s As Double, ByVal a As Double, ByVal b As Double, _
                        ByVal c As Double, ByVal d As Double, ByVal hasD As Boolean, _
                        ByVal mflag As Boolean) As Boolean
    Dim lo As Double, hi As Double, t As Double
    lo = (3# * a + b) / 4#: hi = b
    If lo > hi Then t = lo: lo = hi: hi = t
    If s <= lo Or s >= hi Then Accept = False: Exit Function
    If mflag And Abs(s - b) >= Abs(b - c) / 2# Then Accept = False: Exit Function
    If (Not mflag) And hasD And Abs(s - b) >= Abs(c - d) / 2# Then Accept = False: Exit Function
    If mflag And Abs(b - c) < TOL Then Accept = False: Exit Function
    If (Not mflag) And hasD And Abs(c - d) < TOL Then Accept = False: Exit Function
    Accept = True
End Function


'--- par rate for a gap year: linear in year number between bracketing quotes ---
Private Function InterpYear(ByVal yr As Double, ByRef qYr() As Double, _
                            ByRef qRate() As Double, ByVal nq As Long) As Double
    Dim i As Long, loI As Long, hiI As Long
    loI = 0: hiI = 0
    For i = 1 To nq
        If qYr(i) >= 1# And qYr(i) = Int(qYr(i)) Then
            If qYr(i) <= yr Then
                If loI = 0 Then loI = i ElseIf qYr(i) > qYr(loI) Then loI = i
            End If
            If qYr(i) >= yr Then
                If hiI = 0 Then hiI = i ElseIf qYr(i) < qYr(hiI) Then hiI = i
            End If
        End If
    Next i
    If loI = 0 Then InterpYear = qRate(hiI): Exit Function
    If hiI = 0 Then InterpYear = qRate(loI): Exit Function
    If qYr(hiI) = qYr(loI) Then InterpYear = qRate(loI): Exit Function
    InterpYear = qRate(loI) + (qRate(hiI) - qRate(loI)) * _
                 (yr - qYr(loI)) / (qYr(hiI) - qYr(loI))
End Function


'--- tenor label -> years ------------------------------------------------------
Public Function TenorYears(ByVal tenor As String) As Double
    Dim t As String, u As String, n As Double
    t = UCase$(Trim$(tenor))
    If Len(t) < 2 Then TenorYears = -1: Exit Function
    u = Right$(t, 1)
    If Not IsNumeric(Left$(t, Len(t) - 1)) Then TenorYears = -1: Exit Function
    n = CDbl(Left$(t, Len(t) - 1))
    Select Case u
        Case "D": TenorYears = n / 365#
        Case "W": TenorYears = n * 7# / 365#
        Case "M": TenorYears = n / 12#
        Case "Y": TenorYears = n
        Case Else: TenorYears = -1
    End Select
End Function


'==============================================================================
' TenorDate - maturity for a tenor label off a curve date
'   spot = curve date + 2 business days; W adds 7n days; M/Y add months;
'   then modified following, weekends only.
' Reproduces all 32 dates on the S490 07/21/26 capture.
' (Same routine as CDSBrent.TenorDate - duplicated so this module stands alone.)
'==============================================================================
Public Function TenorDate(ByVal curveDate As Date, ByVal tenor As String) As Variant
    Dim t As String, u As String, n As Long, spot As Date, d As Date
    t = UCase$(Trim$(tenor))
    If Len(t) < 2 Then TenorDate = CVErr(xlErrValue): Exit Function
    u = Right$(t, 1)
    If Not IsNumeric(Left$(t, Len(t) - 1)) Then TenorDate = CVErr(xlErrValue): Exit Function
    n = CLng(Left$(t, Len(t) - 1))
    spot = AddBusDays(curveDate, 2)
    Select Case u
        Case "D": d = spot + n
        Case "W": d = spot + 7 * n
        Case "M": d = DateAdd("m", n, spot)
        Case "Y": d = DateAdd("m", 12 * n, spot)
        Case Else: TenorDate = CVErr(xlErrValue): Exit Function
    End Select
    TenorDate = ModFollowing(d)
End Function


Public Function AddBusDays(ByVal d As Date, ByVal n As Long) As Date
    Dim x As Date
    x = d
    Do While n > 0
        x = x + 1
        If Weekday(x, vbMonday) <= 5 Then n = n - 1
    Loop
    AddBusDays = x
End Function


Private Function ModFollowing(ByVal d As Date) As Date
    Dim x As Date
    x = d
    Do While Weekday(x, vbMonday) > 5
        x = x + 1
    Loop
    If Month(x) <> Month(d) Then
        x = d
        Do While Weekday(x, vbMonday) > 5
            x = x - 1
        Loop
    End If
    ModFollowing = x
End Function
