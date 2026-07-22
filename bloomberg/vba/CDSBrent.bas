Attribute VB_Name = "CDSBrent"
'==============================================================================
' CDSBrent - Brent root-finding for the CDS hazard strip
'
' Mirrors the in-cell bisection on Hazard_Solver exactly: same objective, same
' leg conventions, same inputs. The only difference is the root-finder.
'
'   objective   f(h) = (1-R)*ProtCum(h) - S*( RPV01Cum(h) - D0*D(Ts) )
'               i.e. B-Model (3.3), the par-spread condition
'
'   ProtCum  = ProtFixed  + SUM  Z(mid) * (Q(t-1) - Q(t))
'   RPV01Cum = RPV01Fixed + SUM  alpha * Z(end) * Q(t)
'                         + 1/2 SUM  alpha * Z(mid) * (Q(t-1) - Q(t))
'   Q(t)     = Qstart * exp( -h * cumDt(t) )
'
' Z(mid) on the protection and accrual terms is deliberate: (3.5)/(3.6) integrate
' D(t) at the DEFAULT time, so the midpoint is the second-order approximation.
' Using the payment-date DF instead is ~140x worse against Bloomberg. See
' CLAUDE.md D5 before "fixing" it.
'
' Brent = inverse quadratic interpolation, falling back to secant, falling back
' to bisection when a step would leave the bracket or stop shrinking it. It keeps
' bisection's guarantee of convergence while usually converging much faster.
'
' Import: VBA editor (Alt+F11) > File > Import File > this .bas.
' The workbook must be saved as .xlsm for the module to persist.
'==============================================================================
Option Explicit

Private Const DEFAULT_TOL As Double = 0.000000000000001   ' 1e-15
Private Const DEFAULT_MAXIT As Long = 100
Private Const H_LO As Double = 0#
Private Const H_HI As Double = 3#

' Iterations used by the last CDS_Hazard call - read with CDS_LastIterations()
Private mLastIter As Long


'--- the objective, exposed so f(h) can be plotted or checked at a point --------
Public Function CDS_Objective( _
        ByVal hazard As Double, _
        ByVal marketSpread As Double, _
        ByVal recovery As Double, _
        ByVal qStart As Double, _
        ByVal protFixed As Double, _
        ByVal rpv01Fixed As Double, _
        ByVal aiNet As Double, _
        ByVal alphas As Range, _
        ByVal cumDts As Range, _
        ByVal dfEnd As Range, _
        ByVal dfMid As Range) As Double

    Dim i As Long, n As Long
    Dim a As Double, cum As Double, ze As Double, zm As Double
    Dim q As Double, qPrev As Double, dPD As Double
    Dim prot As Double, rpv As Double

    n = alphas.Cells.Count
    prot = protFixed
    rpv = rpv01Fixed
    qPrev = qStart

    For i = 1 To n
        a = SafeNum(alphas.Cells(i).Value)
        If a <> 0 Then
            cum = SafeNum(cumDts.Cells(i).Value)
            ze = SafeNum(dfEnd.Cells(i).Value)
            zm = SafeNum(dfMid.Cells(i).Value)

            q = qStart * Exp(-hazard * cum)
            dPD = qPrev - q

            prot = prot + zm * dPD
            rpv = rpv + a * ze * q + 0.5 * a * zm * dPD
            qPrev = q
        End If
    Next i

    CDS_Objective = (1# - recovery) * prot - marketSpread * (rpv - aiNet)
End Function


'--- solve for the segment hazard by Brent -------------------------------------
Public Function CDS_Hazard( _
        ByVal marketSpread As Double, _
        ByVal recovery As Double, _
        ByVal qStart As Double, _
        ByVal protFixed As Double, _
        ByVal rpv01Fixed As Double, _
        ByVal aiNet As Double, _
        ByVal alphas As Range, _
        ByVal cumDts As Range, _
        ByVal dfEnd As Range, _
        ByVal dfMid As Range, _
        Optional ByVal tol As Double = DEFAULT_TOL, _
        Optional ByVal maxIter As Long = DEFAULT_MAXIT) As Variant

    Dim a As Double, b As Double, c As Double, d As Double, s As Double
    Dim fa As Double, fb As Double, fc As Double, fs As Double
    Dim mflag As Boolean, it As Long, tmp As Double
    Dim hasD As Boolean

    mLastIter = 0
    a = H_LO
    b = H_HI

    fa = CDS_Objective(a, marketSpread, recovery, qStart, protFixed, rpv01Fixed, _
                       aiNet, alphas, cumDts, dfEnd, dfMid)
    fb = CDS_Objective(b, marketSpread, recovery, qStart, protFixed, rpv01Fixed, _
                       aiNet, alphas, cumDts, dfEnd, dfMid)

    ' No sign change means the quote is unreachable on [0,3]. That is the
    ' strippability failure of white paper p.9, usually an inverted curve.
    ' Return an error rather than a bracket end pretending to be a root.
    If fa * fb > 0# Then
        CDS_Hazard = CVErr(xlErrNum)
        Exit Function
    End If

    If Abs(fa) < Abs(fb) Then
        tmp = a: a = b: b = tmp
        tmp = fa: fa = fb: fb = tmp
    End If

    c = a: fc = fa
    mflag = True
    hasD = False
    d = 0#

    Do While fb <> 0# And Abs(b - a) > tol
        it = it + 1
        If it > maxIter Then Exit Do

        If (fa <> fc) And (fb <> fc) Then
            ' inverse quadratic interpolation
            s = a * fb * fc / ((fa - fb) * (fa - fc)) _
              + b * fa * fc / ((fb - fa) * (fb - fc)) _
              + c * fa * fb / ((fc - fa) * (fc - fb))
        Else
            ' secant
            s = b - fb * (b - a) / (fb - fa)
        End If

        If Not BrentAccept(s, a, b, c, d, hasD, mflag, tol) Then
            s = (a + b) / 2#
            mflag = True
        Else
            mflag = False
        End If

        fs = CDS_Objective(s, marketSpread, recovery, qStart, protFixed, _
                           rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)

        d = c: hasD = True
        c = b: fc = fb

        If fa * fs < 0# Then
            b = s: fb = fs
        Else
            a = s: fa = fs
        End If

        If Abs(fa) < Abs(fb) Then
            tmp = a: a = b: b = tmp
            tmp = fa: fa = fb: fb = tmp
        End If
    Loop

    mLastIter = it
    CDS_Hazard = b
End Function


'--- model par spread in bp at a given hazard ----------------------------------
Public Function CDS_ModelSpread( _
        ByVal hazard As Double, _
        ByVal recovery As Double, _
        ByVal qStart As Double, _
        ByVal protFixed As Double, _
        ByVal rpv01Fixed As Double, _
        ByVal aiNet As Double, _
        ByVal alphas As Range, _
        ByVal cumDts As Range, _
        ByVal dfEnd As Range, _
        ByVal dfMid As Range) As Variant

    Dim prot As Double, rpv As Double
    ' f(h) = (1-R)*Prot - S*(RPV01 - aiNet); at S = 0 the objective IS (1-R)*Prot
    prot = CDS_Objective(hazard, 0#, recovery, qStart, protFixed, rpv01Fixed, _
                         aiNet, alphas, cumDts, dfEnd, dfMid)
    ' recover RPV01 by differencing at S = 1
    rpv = prot - CDS_Objective(hazard, 1#, recovery, qStart, protFixed, _
                               rpv01Fixed, aiNet, alphas, cumDts, dfEnd, dfMid)

    If rpv = 0# Then
        CDS_ModelSpread = CVErr(xlErrDiv0)
    Else
        CDS_ModelSpread = prot / rpv * 10000#
    End If
End Function


'--- iterations used by the last CDS_Hazard call --------------------------------
Public Function CDS_LastIterations() As Long
    CDS_LastIterations = mLastIter
End Function


'==============================================================================
' helpers
'==============================================================================
Private Function BrentAccept(ByVal s As Double, ByVal a As Double, ByVal b As Double, _
                             ByVal c As Double, ByVal d As Double, ByVal hasD As Boolean, _
                             ByVal mflag As Boolean, ByVal tol As Double) As Boolean
    Dim lo As Double, hi As Double
    lo = (3# * a + b) / 4#
    hi = b
    If lo > hi Then
        Dim t As Double
        t = lo: lo = hi: hi = t
    End If

    If s <= lo Or s >= hi Then BrentAccept = False: Exit Function
    If mflag And Abs(s - b) >= Abs(b - c) / 2# Then BrentAccept = False: Exit Function
    If (Not mflag) And hasD And Abs(s - b) >= Abs(c - d) / 2# Then BrentAccept = False: Exit Function
    If mflag And Abs(b - c) < tol Then BrentAccept = False: Exit Function
    If (Not mflag) And hasD And Abs(c - d) < tol Then BrentAccept = False: Exit Function

    BrentAccept = True
End Function


Private Function SafeNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = 0#
    End If
End Function


'==============================================================================
' TenorDate - maturity for a tenor label off a curve date
'   spot   = curve date + 2 business days
'   1W..3W = spot + 7*n days
'   nM     = spot + n months
'   nY     = spot + 12n months
'   then modified following, weekends only (no holiday calendar)
' Reproduces all 32 dates on the S490 07/21/26 capture.
'==============================================================================
Public Function TenorDate(ByVal curveDate As Date, ByVal tenor As String) As Variant
    Dim t As String, unit As String, n As Long, spot As Date, d As Date
    t = UCase$(Trim$(tenor))
    If Len(t) < 2 Then TenorDate = CVErr(xlErrValue): Exit Function

    unit = Right$(t, 1)
    If Not IsNumeric(Left$(t, Len(t) - 1)) Then TenorDate = CVErr(xlErrValue): Exit Function
    n = CLng(Left$(t, Len(t) - 1))

    spot = AddBusDays(curveDate, 2)
    Select Case unit
        Case "W": d = spot + 7 * n
        Case "M": d = DateAdd("m", n, spot)
        Case "Y": d = DateAdd("m", 12 * n, spot)
        Case Else: TenorDate = CVErr(xlErrValue): Exit Function
    End Select

    TenorDate = ModFollowing(d)
End Function


Private Function AddBusDays(ByVal d As Date, ByVal n As Long) As Date
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

'==============================================================================
' CDS_StripHazard - the whole strip in one call, straight off CDS_Schedule.
'
' Hazard_Solver existed to stage inputs: it sliced CDS_Schedule into per-tenor
' row vectors for CDS_Hazard to read, and held a bisection ladder as the solver.
' This does the slicing internally, so the Brent path no longer touches
' Hazard_Solver at all.
'
'   k          which tenor, 1..N
'   mats       maturity date per tenor        Hazard_Bootstrap!$B$7:$B$12
'   spreads    market spread in bp per tenor  Hazard_Bootstrap!$C$7:$C$12
'   recovery   R                              CDS_Parameters!$B$8
'   aiNet      D0 * D(Ts)                     CDS_Parameters!$B$26
'   payDates   pay date per period            CDS_Schedule!$C$7:$C$46
'   alphas     accrual ACT/360                CDS_Schedule!$D$7:$D$46
'   dts        dt ACT/365                     CDS_Schedule!$E$7:$E$46
'   dfEnd      D(pay date)                    CDS_Schedule!$G$7:$G$46
'   dfMid      D(period midpoint)             CDS_Schedule!$H$7:$H$46
'
' No circularity: D, E, G and H are accruals and discount factors. None of them
' depend on the hazard column, so passing those ranges back in is safe. Do NOT
' pass column I - that one does depend on the output.
'
' Section 4, p.8: solve maturity by maturity, holding earlier hazards fixed.
' The whole strip is computed on the first call and cached, so six cells cost
' one strip rather than six.
'==============================================================================
Private sKey As String
Private sH(1 To 32) As Double
Private sN As Long

Public Function CDS_StripHazard(ByVal k As Long, ByVal mats As Range, ByVal spreads As Range, _
        ByVal recovery As Double, ByVal aiNet As Double, ByVal payDates As Range, _
        ByVal alphas As Range, ByVal dts As Range, ByVal dfEnd As Range, _
        ByVal dfMid As Range) As Variant

    Dim nT As Long, nP As Long, i As Long, j As Long
    Dim key As String, S As Double, h As Double
    Dim qPrev As Double, protFix As Double, rpvFix As Double
    Dim i0 As Long, iLo As Long, iHi As Long
    Dim q As Double, qn As Double, d As Double

    nT = mats.Cells.Count
    nP = alphas.Cells.Count
    If k < 1 Or k > nT Then CDS_StripHazard = CVErr(xlErrValue): Exit Function

    key = CStr(recovery) & "|" & CStr(aiNet) & "|"
    For i = 1 To nT
        key = key & CStr(NumOf(mats.Cells(i).Value)) & ":" & CStr(NumOf(spreads.Cells(i).Value)) & ";"
    Next i
    For i = 1 To nP
        key = key & Format$(NumOf(dfEnd.Cells(i).Value), "0.000000000") & ","
    Next i

    If key <> sKey Then
        qPrev = 1#: protFix = 0#: rpvFix = 0#: i0 = 1
        For j = 1 To nT
            iLo = i0: iHi = i0 - 1
            For i = i0 To nP
                If NumOf(alphas.Cells(i).Value) > 0# And _
                   NumOf(payDates.Cells(i).Value) <= NumOf(mats.Cells(j).Value) + 0.5 Then
                    iHi = i
                End If
            Next i
            S = NumOf(spreads.Cells(j).Value) / 10000#
            h = SolveSeg(S, recovery, aiNet, qPrev, protFix, rpvFix, iLo, iHi, _
                         alphas, dts, dfEnd, dfMid)
            sH(j) = h
            q = qPrev
            For i = iLo To iHi
                If NumOf(alphas.Cells(i).Value) > 0# Then
                    qn = q * Exp(-h * NumOf(dts.Cells(i).Value))
                    d = q - qn
                    protFix = protFix + NumOf(dfMid.Cells(i).Value) * d
                    rpvFix = rpvFix + NumOf(alphas.Cells(i).Value) * NumOf(dfEnd.Cells(i).Value) * qn _
                           + 0.5 * NumOf(alphas.Cells(i).Value) * NumOf(dfMid.Cells(i).Value) * d
                    q = qn
                End If
            Next i
            qPrev = q
            If iHi >= iLo Then i0 = iHi + 1
        Next j
        sKey = key: sN = nT
    End If

    CDS_StripHazard = sH(k)
End Function


' one segment by Brent, earlier hazards already baked into protFix / rpvFix
Private Function SolveSeg(ByVal S As Double, ByVal R As Double, ByVal aiNet As Double, _
        ByVal qStart As Double, ByVal protFix As Double, ByVal rpvFix As Double, _
        ByVal iLo As Long, ByVal iHi As Long, ByVal alphas As Range, ByVal dts As Range, _
        ByVal dfEnd As Range, ByVal dfMid As Range) As Double

    Dim a As Double, b As Double, c As Double, dd As Double, s2 As Double
    Dim fa As Double, fb As Double, fc As Double, fs As Double
    Dim mflag As Boolean, hasD As Boolean, it As Long, t As Double

    a = 0#: b = 3#
    fa = SegF(a, S, R, aiNet, qStart, protFix, rpvFix, iLo, iHi, alphas, dts, dfEnd, dfMid)
    fb = SegF(b, S, R, aiNet, qStart, protFix, rpvFix, iLo, iHi, alphas, dts, dfEnd, dfMid)
    If fa * fb > 0# Then SolveSeg = 0#: Exit Function      ' p.9 not strippable

    If Abs(fa) < Abs(fb) Then
        t = a: a = b: b = t
        t = fa: fa = fb: fb = t
    End If
    c = a: fc = fa: mflag = True: hasD = False: dd = 0#

    Do While fb <> 0# And Abs(b - a) > 0.000000000000001 And it < 200
        it = it + 1
        If (fa <> fc) And (fb <> fc) Then
            s2 = a * fb * fc / ((fa - fb) * (fa - fc)) _
               + b * fa * fc / ((fb - fa) * (fb - fc)) _
               + c * fa * fb / ((fc - fa) * (fc - fb))
        Else
            s2 = b - fb * (b - a) / (fb - fa)
        End If
        If Not SegAcc(s2, a, b, c, dd, hasD, mflag) Then
            s2 = (a + b) / 2#: mflag = True
        Else
            mflag = False
        End If
        fs = SegF(s2, S, R, aiNet, qStart, protFix, rpvFix, iLo, iHi, alphas, dts, dfEnd, dfMid)
        dd = c: hasD = True
        c = b: fc = fb
        If fa * fs < 0# Then b = s2: fb = fs Else a = s2: fa = fs
        If Abs(fa) < Abs(fb) Then
            t = a: a = b: b = t
            t = fa: fa = fb: fb = t
        End If
    Loop
    SolveSeg = b
End Function


Private Function SegF(ByVal hz As Double, ByVal S As Double, ByVal R As Double, _
        ByVal aiNet As Double, ByVal qStart As Double, ByVal protFix As Double, _
        ByVal rpvFix As Double, ByVal iLo As Long, ByVal iHi As Long, _
        ByVal alphas As Range, ByVal dts As Range, ByVal dfEnd As Range, _
        ByVal dfMid As Range) As Double
    Dim i As Long, q As Double, qn As Double, d As Double, a As Double
    Dim prot As Double, rpv As Double
    prot = protFix: rpv = rpvFix: q = qStart
    For i = iLo To iHi
        a = NumOf(alphas.Cells(i).Value)
        If a > 0# Then
            qn = q * Exp(-hz * NumOf(dts.Cells(i).Value))
            d = q - qn
            prot = prot + NumOf(dfMid.Cells(i).Value) * d
            rpv = rpv + a * NumOf(dfEnd.Cells(i).Value) * qn + 0.5 * a * NumOf(dfMid.Cells(i).Value) * d
            q = qn
        End If
    Next i
    SegF = (1# - R) * prot - S * (rpv - aiNet)
End Function


Private Function SegAcc(ByVal s As Double, ByVal a As Double, ByVal b As Double, _
                        ByVal c As Double, ByVal d As Double, ByVal hasD As Boolean, _
                        ByVal mflag As Boolean) As Boolean
    Dim lo As Double, hi As Double, t As Double
    lo = (3# * a + b) / 4#: hi = b
    If lo > hi Then t = lo: lo = hi: hi = t
    If s <= lo Or s >= hi Then SegAcc = False: Exit Function
    If mflag And Abs(s - b) >= Abs(b - c) / 2# Then SegAcc = False: Exit Function
    If (Not mflag) And hasD And Abs(s - b) >= Abs(c - d) / 2# Then SegAcc = False: Exit Function
    SegAcc = True
End Function


Private Function NumOf(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = 0#
End Function

