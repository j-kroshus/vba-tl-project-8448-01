'* Data Collection TRANSFER Button (Sheet2)
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21  Original Sub BTN_TRANSFER()
Sub BTN_TRANSFER()

  Dim job_Arr() As String: job_Arr() = IMPORT_JOB_DATA()
  Dim pcd_Arr() As String: pcd_Arr() = IMPORT_PCD_DATA()
  
  Call IMPORT_WS_COL(pcd_Arr)
  
End Sub

'* Data Collection ACCEPT Button (Sheet2)
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21 Original Sub BTN_ACCEPT()
'*** TESTED *** Author: Josh Kroshus Updated: 11-14-21 Updated Logic In Select Case
'*** TESTED *** Author: Josh Kroshus Updated: 11-19-21 Updated Logic In Select Case For EMPTY CELLS
'*** TESTED *** Author: Josh Kroshus Updated: 11-19-21 Updated Logic In Select Case For FP, CC, LP
'*** TESTED *** Author: Josh Kroshus Updated: 11-22-21 Updated Logic In Select Case For ADJUST, RESET, AND VERIFY
Sub BTN_ACCEPT()

  Dim job_Arr() As String: job_Arr() = IMPORT_JOB_DATA()
  Dim pcd_Arr() As String: pcd_Arr() = IMPORT_PCD_DATA()
  
  Dim N2 As String: N2 = ThisWorkbook.Worksheets("DATA COLLECTION").Range("N2").Value
  Dim O2 As String: O2 = ThisWorkbook.Worksheets("DATA COLLECTION").Range("O2").Value
  
  Dim WS As Worksheet: Set WS = ThisWorkbook.Worksheets("DATA COLLECTION")

  '* <><><><><> SAMPLE PASSED <><><><><>
  Select Case True

    '* <><><><><> IS JOB # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("D2").Value) = True
      MsgBox "Job# Is A Required Field", vbCritical, "ERROR:"
      Exit Sub

    '* <><><><><> IS HEAT # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("F2").Value) = True
      MsgBox "Heat# Is A Required Field", vbCritical + vbYesNo, "ERROR:"
      Exit Sub

    '* <><><><><> IS COIL # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("H2").Value) = True
      MsgBox "Coil# Is A Required Field", vbCritical + vbYesNo, "ERROR:"
      Exit Sub

    '* <><><><><> A, B, C, C-1 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" ) And ( O2 = "FP" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1") And (O2 = "FP")
      If MsgBox("Is Job Information Correct?", vbQuestion + vbYesNo, "NOTICE:") = vbNo Then
        MsgBox "Update The Job Information Before Continuing"
        End If
      Exit Sub

    '* <><><><><> A, B, C, C-1 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" ) And ( O2 = "CC" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1") And (O2 = "CC")
      If MsgBox("Is Lot# and Coil# Correct?", vbQuestion + vbYesNo, "NOTICE:") = vbNo Then
        WS.Range("F2") = InputBox("Enter The Correct Lot# Now.", "Lot# ERROR:")
        WS.Range("H2") = InputBox("Enter The Correct Coil# Now.", "Coil# ERROR:")
        End If
      job_Arr(1) = "PASS"
      job_Arr(3) = WS.Range("F2").Value 'HEAT
      job_Arr(4) = WS.Range("H2").Value 'COIL
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
      Call IMPORT_WS_CPK(job_Arr, pcd_Arr)
      
    '* <><><><><> A, B, C, C-2 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-2" ) And ( O2 = "LP" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-2") And (O2 = "LP")
      If MsgBox("You Are About To Save Job" & WS.Range("D2") & " " & Format(Now(), "MM-DD-YY") & ".xls" _
        & vbCrLf & "To The Inspection History Folder, Do You Want To Continue?", vbQuestion + vbYesNo, "NOTICE:") = vbYes Then
        End If
      job_Arr(1) = "PASS"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
      Call IMPORT_WS_CPK(job_Arr, pcd_Arr)
      Call WS_DELETE("PCDmisExcel7")
      Call BTN_CLOSE
      Exit Sub

    '* <><><><><> A, B, C, C-1 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" ) And ( O2 = "ADJUST" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1") And (O2 = "ADJUST")
      job_Arr(1) = "PASS"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)

    '* <><><><><> A, B, C, C-1 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" ) And ( O2 = "RESET" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1") And (O2 = "RESET")
      job_Arr(1) = "PASS"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)

    '* <><><><><> A, B, C, C-1 DIE <><><><><> ( N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" ) And ( O2 = "VERIFY" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1") And (O2 = "VERIFY")
      job_Arr(1) = "PASS"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)

    Case Else
      job_Arr(1) = "PASS"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
      Call IMPORT_WS_CPK(job_Arr, pcd_Arr)

  End Select

  Call WS_DELETE("PCDmisExcel7")
  Call WS_CLEAR_COL
  Call WB_CLOSE

End Sub

'* Data Collection REJECT Button (Sheet2)
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21 Original Sub BTN_REJECT()
'*** TESTED *** Author: Josh Kroshus Updated: 11-14-21 Updated Logic In Select Case
'*** TESTED *** Author: Josh Kroshus Updated: 11-19-21 Updated Logic In Select Case For EMPTY CELLS
'*** TESTED *** Author: Josh Kroshus Updated: 11-22-21 Updated Logic In Select Case Removed Logic For FP, CC, LP Not Needed For Sample Rejects
'*** TESTED *** Author: Josh Kroshus Updated: 11-22-21 Updated Logic In Select Case For ADJUST, RESET, AND VERIFY
'! NEED TO ADD COMMENTS FIELD REQUIRED LOGIC FOR REJECTED SAMPLES
Sub BTN_REJECT()

  Dim job_Arr() As String: job_Arr() = IMPORT_JOB_DATA()
  Dim pcd_Arr() As String: pcd_Arr() = IMPORT_PCD_DATA()
  
  Dim N2 As String: N2 = ThisWorkbook.Worksheets("DATA COLLECTION").Range("N2").Value
  Dim O2 As String: O2 = ThisWorkbook.Worksheets("DATA COLLECTION").Range("O2").Value
  
  Dim WS As Worksheet: Set WS = ThisWorkbook.Worksheets("DATA COLLECTION")

  '* <><><><><> SAMPLE FAILED <><><><><>
  Select Case True

    '* <><><><><> IS JOB # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("D2").Value) = True
      MsgBox "Job # Is A Required Field", vbCritical, "ERROR:"
      Exit Sub

    '* <><><><><> IS HEAT # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("F2").Value) = True
      MsgBox "Heat # Is A Required Field", vbCritical + vbYesNo, "ERROR:"
      Exit Sub

    '* <><><><><> IS COIL # EMPTY LOGIC <><><><><>
    Case IsEmpty(Range("H2").Value) = True
      MsgBox "Coil # Is A Required Field", vbCritical + vbYesNo, "ERROR:"
      Exit Sub

    '* <><><><><> A, B, C, C-1, C-2 DIE <><><><><> (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And ( O2 = "ADJUST" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And (O2 = "ADJUST")
      job_Arr(1) = "FAIL"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)

    '* <><><><><> A, B, C, C-1, C-2 DIE <><><><><> (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And ( O2 = "RESET" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And (O2 = "RESET")
      job_Arr(1) = "FAIL"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)

    '* <><><><><> A, B, C, C-1, C-2 DIE <><><><><> (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And ( O2 = "VERIFY" )
    Case (N2 = "A" Or N2 = "B" Or N2 = "C" Or N2 = "C-1" Or N2 = "C-2") And (O2 = "VERIFY")
      job_Arr(1) = "FAIL"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
      
    Case Else
      job_Arr(1) = "FAIL"
      job_Arr(12) = Format(Now, "MM-DD-YY")
      Call IMPORT_WS_HIS(job_Arr, pcd_Arr)
      
  End Select

  Call WS_DELETE("PCDmisExcel7")
  Call WS_CLEAR_COL
  Call WB_CLOSE

End Sub

'* Data Collection CLOSE Button (Sheet2)
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21
'*** TESTED *** Author: Josh Kroshus Updated: 11-09-21
Sub BTN_CLOSE()

  Dim lRow, lCol As Integer
  Dim fPath As String: fPath = "Q:\IQS LINKS\Product\Trueline Produced Parts\8448\Current Revision\Production Inspection History\"
  Dim fName As String: fName = Worksheets("DATA COLLECTION").Range("D2").Text
  Dim fDate As String: fDate = Format(Now(), "MM-DD-YY")

  ThisWorkbook.Save
  ThisWorkbook.SaveCopyAs Filename:=fPath & fName & " " & fDate & ".xls"

  Call WS_CLEAR_COL
  Call WS_CLEAR("DATA HISTORY")
  Call WS_CLEAR("CPK")
  Call WB_CLOSE

End Sub

