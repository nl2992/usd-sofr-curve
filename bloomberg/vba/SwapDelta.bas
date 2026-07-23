Attribute VB_Name = "SwapDelta"
'==============================================================================
' SwapDelta - bucketed key-rate delta for the amortising swap, by bump & reprice.
'
' Same idea as the CDS risk: nudge the curve, let the workbook re-bootstrap, read
' the new NPV, take the difference. The whole chain Quotes -> Interpolation ->
' Bootstrap -> Amort_Swap recalculates on any quote change, so this Sub only has
' to bump one quote at a time and read Amort_Swap!B17.
'
'   key-rate delta(tenor) = NPV(that tenor +1bp) - NPV(base)
'   total DV01            = NPV(all tenors +1bp) - NPV(base)
'
' The trade fixed rate is FROZEN first (B12 <- value of the solved par B11), so
' it does not re-solve to par under each bump - otherwise every NPV would stay 0
' and every delta would be zero.
'
' Run:  Alt+F8 > RunSwapDelta.  Results land on the Delta sheet.
' Needs the file saved as .xlsm.
'==============================================================================
Option Explicit

Public Sub RunSwapDelta()
    Dim wq As Worksheet, wa As Worksheet, wd As Worksheet
    Dim r As Long, n As Long, outRow As Long
    Dim base As Double, saved As Double, d As Double, sumBuckets As Double
    Dim savedAll() As Double
    Const BUMP As Double = 0.0001          ' 1bp as a fraction

    Set wq = ThisWorkbook.Worksheets("Quotes")
    Set wa = ThisWorkbook.Worksheets("Amort_Swap")
    On Error Resume Next
    Set wd = ThisWorkbook.Worksheets("Delta")
    On Error GoTo 0
    If wd Is Nothing Then Set wd = ThisWorkbook.Worksheets.Add(After:=wa): wd.Name = "Delta"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ' freeze the trade fixed rate at the solved par, so it does not chase the bump
    wa.Range("B12").Value = wa.Range("B11").Value
    Application.Calculate
    base = wa.Range("B17").Value

    ' header
    wd.Cells.Clear
    wd.Range("A1").Value = "Bucketed key-rate delta (+1bp per tenor)"
    wd.Range("A1").Font.Bold = True
    wd.Range("A2").Value = "Trade frozen at " & Format(wa.Range("B12").Value, "0.0000%") & _
                           ".  Base NPV " & Format(base, "#,##0")
    wd.Range("A4").Value = "Tenor": wd.Range("B4").Value = "Delta (+1bp, KRW)"
    wd.Range("A4:B4").Font.Bold = True

    ' quotes live in Quotes!B5:B19 (year in C, label in A)
    outRow = 5: sumBuckets = 0
    For r = 5 To 19
        If IsNumeric(wq.Cells(r, 2).Value) And wq.Cells(r, 2).Value <> 0 Then
            saved = wq.Cells(r, 2).Value
            wq.Cells(r, 2).Value = saved + BUMP
            Application.Calculate
            d = wa.Range("B17").Value - base
            wq.Cells(r, 2).Value = saved
            wd.Cells(outRow, 1).Value = wq.Cells(r, 1).Value   ' tenor label
            wd.Cells(outRow, 2).Value = d
            wd.Cells(outRow, 2).NumberFormat = "#,##0"
            sumBuckets = sumBuckets + d
            outRow = outRow + 1
        End If
    Next r
    Application.Calculate

    ' total DV01 - bump every quote at once
    ReDim savedAll(5 To 19)
    For r = 5 To 19
        savedAll(r) = wq.Cells(r, 2).Value
        If IsNumeric(savedAll(r)) And savedAll(r) <> 0 Then wq.Cells(r, 2).Value = savedAll(r) + BUMP
    Next r
    Application.Calculate
    Dim dv01 As Double: dv01 = wa.Range("B17").Value - base
    For r = 5 To 19
        wq.Cells(r, 2).Value = savedAll(r)
    Next r
    Application.Calculate

    outRow = outRow + 1
    wd.Cells(outRow, 1).Value = "Sum of buckets": wd.Cells(outRow, 1).Font.Bold = True
    wd.Cells(outRow, 2).Value = sumBuckets: wd.Cells(outRow, 2).NumberFormat = "#,##0"
    wd.Cells(outRow + 1, 1).Value = "Total DV01 (parallel)": wd.Cells(outRow + 1, 1).Font.Bold = True
    wd.Cells(outRow + 1, 2).Value = dv01: wd.Cells(outRow + 1, 2).NumberFormat = "#,##0"
    wd.Cells(outRow + 3, 1).Value = "Buckets should sum to ~DV01; small gap is convexity."
    wd.Columns("A:B").AutoFit

    Application.ScreenUpdating = True
    MsgBox "Delta done. DV01 " & Format(dv01, "#,##0") & " KRW/bp.", vbInformation
End Sub
