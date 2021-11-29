'* JOSH KROSHUS DATE: 11-04-21
'* PCDMIS WORKSHEET LOGIC (FUTURE UPDATES)
'! <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>
If WS = PCDmisExcelA Then


If WS = PCDmisExcelB Then


If WS = PCDmisExcelC Then


If WS = PCDmisExcel2C Then


'https://excelmacromastery.com/vba-dim/

$ String
% Integer
& Long
! Single
# Double

Range("A1").Select

With ActiveWindow
    .SplitColumn = col
    .SplitRow = row
    .FreezePanes = True
End With

End Sub

Sub INSPECTION_REPORT_SETUP()

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "DASHBOARD"

Call FREEZE_ROWS(5, 0)

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "CUST DATA"
Call FREEZE_ROWS(5, 0)

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "HIST DATA"
Call FREEZE_ROWS(5, 0)

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "EDIT LOG"
Call FREEZE_ROWS(1, 0)

End Sub

Sub FREEZE_ROWS(row, col)

Private Sub CreateCommandButton()
Dim ctop#, cleft#, cht#, cwdth#
Dim sht As Worksheet
Dim Btn As OLEObject
Set sht = ThisWorkbook.Worksheets("DATA HISTORY")
With Range("A1")
ctop = .Top
cleft = .Left
cht = .Height
cwdth = .Width
End With
With sht
Set Btn = .OLEObjects.Add(ClassType:="Forms.ToggleButton.1", Left:=cleft, Top:=ctop, Width:=cwdth, Height:=cht)
End With
Btn.Object.Caption = "+"
Btn.Name = "TBTN_EXPAND_HIS"
Btn.Placement = xlMoveAndSize
'Optional code insertion - - establish ref in VBE to MS VBA Extensibility 5.3 Library
With ThisWorkbook.VBProject.VBComponents(sht.CodeName).CodeModule
.InsertLines .CreateEventProc("Click", Btn.Name) + 1, "Msgbox ""Replace this message with your actual code."" "
End With
End Sub



    job_Arr(1) = "PASS"
    job_Arr(9) = "LP"
    job_Arr(12) = Format(Now, "MM-DD-YY")
    Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
    Call IMPORT_WS_CPK(job_Arr, pcd_Arr)

      Call BTN_CLOSE()

      WS.Range("D2") = InputBox("Enter The Correct JOB# Now.", "NEW JOB#:")
      WS.Range("F2") = InputBox("Enter The Correct LOT# Now.", "NEW LOT#:")
      WS.Range("H2") = InputBox("Enter The Correct COIL# Now.", "NEW COIL#:")
      WS.Range("N2") = InputBox("Enter The Correct DIE Now.", "NEW DIE:")
      WS.Range("N2") = job_Arr(9) = "FP"
      WS.Range("P2") = InputBox("Enter The Inspector Now.", "NEW DIE:")

    job_Arr(1) = "PASS"
    job_Arr(12) = Format(Now, "MM-DD-YY")
    Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
    Call IMPORT_WS_CPK(job_Arr, pcd_Arr)

      
    '* <><><><><> A, B, C, C-2 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-2" ) And ( O2 = "LP/FP" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-2") And (O2 = "LP/FP")
      If MsgBox("You Are About To Save Job" & WS.Range("D2") & " " & Format(Now(), "MM-DD-YY") & ".xls" _
        & vbCrLf & "To The Inspection History Folder, Do You Want To Continue?", vbQuestion + vbYesNo, "NOTICE:") = vbYes Then
        MsgBox "You have Selected Yes"
        End If