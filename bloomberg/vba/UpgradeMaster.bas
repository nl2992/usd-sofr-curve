Attribute VB_Name = "UpgradeMaster"
'==============================================================================
' UpgradeMaster - run ONCE on the master, then Testing moves in cleanly.
'
' Does the same four things upgrade_master.py did, so no Python is needed:
'
'   1. the ten BOOT_* defined names Testing binds to. Names are the whole trick:
'      a direct cross-sheet reference gets rewritten as a link back to the
'      source file when a sheet is copied, but a name binds to the DESTINATION's
'      definition instead.
'   2. SOFR_OIS_Quotes!H - reads the three Testing blocks, matched on maturity
'      DATE so a 12M label lines up with a 1Y row
'   3. Instructions!B9  - VAL_DATE follows the active test case's curve date
'   4. Bootstrap!G4     - dropdown gains Test 1 / Test 2 / Test 3
'
' HOW TO RUN
'   Alt+F11 > File > Import File > this .bas
'   then in the Immediate window (Ctrl+G):   UpgradeMaster
'   or Alt+F8 > UpgradeMaster > Run
'
' Safe to re-run - it deletes and re-adds rather than duplicating. It does NOT
' touch the Testing sheet itself.
'
' SAVE A COPY OF THE MASTER FIRST. This edits formulas in place and VBA edits
' are not covered by Excel's undo.
'==============================================================================
Option Explicit

Private Const MODE_REF As String = "Bootstrap!$G$4"
Private Const FIXED_LBL As String = "Fixed (S490 07/21/26)"
Private Const LIVE_LBL As String = "Live (BDP)"

Public Sub UpgradeMaster()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long, r As Long, n As Long
    Dim f As String, inner As String
    Dim starts As Variant, ends As Variant, cds As Variant
    Dim nm As Variant, ref As Variant

    Set wb = ActiveWorkbook

    ' the sheets this expects to find
    For Each nm In Array("Bootstrap", "SOFR_OIS_Quotes", "Instructions")
        If Not SheetExists(wb, CStr(nm)) Then
            MsgBox "No '" & nm & "' sheet in " & wb.Name & "." & vbCrLf & _
                   "This is not the master workbook - open that one and run again.", _
                   vbExclamation, "Upgrade aborted"
            Exit Sub
        End If
    Next nm

    If MsgBox("Upgrade '" & wb.Name & "' so the Testing sheet can be moved in?" & vbCrLf & vbCrLf & _
              "This edits formulas in place and cannot be undone." & vbCrLf & _
              "Make sure you have a copy of the file first.", _
              vbOKCancel + vbQuestion, "Upgrade master") <> vbOK Then Exit Sub

    ' ---- 1. defined names -------------------------------------------------
    nm = Array("BOOT_MODE", "BOOT_DFSPOT", "BOOT_DATES", "BOOT_S", "BOOT_TAU0", _
               "BOOT_TAUC", "BOOT_ANN", "BOOT_DF", "BOOT_T", "BOOT_ZERO")
    ref = Array("Bootstrap!$G$4", "Bootstrap!$D$4", "Bootstrap!$B$8:$B$72", _
                "Bootstrap!$E$8:$E$72", "Bootstrap!$D$8:$D$72", "Bootstrap!$F$8:$F$72", _
                "Bootstrap!$G$8:$G$72", "Bootstrap!$H$8:$H$72", "Bootstrap!$C$8:$C$72", _
                "Bootstrap!$J$8:$J$72")
    For i = LBound(nm) To UBound(nm)
        On Error Resume Next
        wb.Names(CStr(nm(i))).Delete
        On Error GoTo 0
        wb.Names.Add Name:=CStr(nm(i)), RefersTo:="=" & CStr(ref(i))
    Next i

    ' ---- 2. SOFR_OIS_Quotes!H --------------------------------------------
    starts = Array(8, 58, 108)
    ends = Array(47, 97, 147)
    Set ws = wb.Worksheets("SOFR_OIS_Quotes")
    n = 0
    For r = 5 To 39
        If Len(Trim$(CStr(ws.Cells(r, "A").Value))) > 0 _
           And Len(Trim$(CStr(ws.Cells(r, "B").Value))) > 0 Then
            f = "=IFERROR("
            For i = 0 To 2
                f = f & "IF(" & MODE_REF & "=""Test " & (i + 1) & """," & _
                    "INDEX(Testing!$B$" & starts(i) & ":$B$" & ends(i) & "," & _
                    "MATCH(C" & r & ",Testing!$F$" & starts(i) & ":$F$" & ends(i) & ",0)),"
            Next i
            inner = "IF(" & MODE_REF & "=""" & FIXED_LBL & """,J" & r & _
                    ",IF(ISNUMBER(T" & r & "),T" & r & ",J" & r & "))"
            f = f & inner & ")))" & ",J" & r & ")"
            ws.Cells(r, "H").Formula = f
            ws.Cells(r, "H").NumberFormat = "0.00000"
            n = n + 1
        End If
    Next r

    ' ---- 3. Instructions!B9 ----------------------------------------------
    cds = Array("Testing!$D$5", "Testing!$D$55", "Testing!$D$105")
    Set ws = wb.Worksheets("Instructions")
    f = "=IF(" & MODE_REF & "=""Test 1""," & cds(0) & _
        ",IF(" & MODE_REF & "=""Test 2""," & cds(1) & _
        ",IF(" & MODE_REF & "=""Test 3""," & cds(2) & ",DATE(2026,7,21))))"
    ws.Range("B9").Formula = f
    ws.Range("B9").NumberFormat = "mm/dd/yyyy"

    ' ---- 4. Bootstrap!G4 dropdown ----------------------------------------
    Set ws = wb.Worksheets("Bootstrap")
    With ws.Range("G4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=LIVE_LBL & "," & FIXED_LBL & ",Test 1,Test 2,Test 3"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    Application.CalculateFullRebuild

    MsgBox "Upgraded '" & wb.Name & "'." & vbCrLf & vbCrLf & _
           "  " & (UBound(nm) - LBound(nm) + 1) & " defined names" & vbCrLf & _
           "  " & n & " quote rows repointed at Testing (matched on date)" & vbCrLf & _
           "  VAL_DATE now follows the active test case" & vbCrLf & _
           "  Bootstrap!G4 dropdown has Test 1-3" & vbCrLf & vbCrLf & _
           "From now on: right-click Testing > Move or Copy > into this workbook," & vbCrLf & _
           "delete the old Testing, rename the new one to exactly 'Testing'." & vbCrLf & vbCrLf & _
           "Test 2 calls VBA - import CurveVBA.bas here too and save as .xlsm.", _
           vbInformation, "Done"
End Sub


Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function
