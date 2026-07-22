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
