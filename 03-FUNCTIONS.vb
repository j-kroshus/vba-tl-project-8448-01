'* 8448 ARRAY FUNCTIONS
'* <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>

'* Write Data To JOB_ARR
'*** TESTED *** Name: Josh Kroshus Date: 10-19-21
Function IMPORT_JOB_DATA()

  Dim WS As Worksheet: Set WS = Worksheets("DATA COLLECTION")
  Dim job_Arr(15) As String

  job_Arr(2) = WS.Range("D2").Value 'JOB
  job_Arr(3) = WS.Range("F2").Value 'HEAT
  job_Arr(4) = WS.Range("H2").Value 'COIL
  job_Arr(5) = WS.Range("J2").Value 'TAPE
  job_Arr(6) = WS.Range("L2").Value 'PART
  job_Arr(7) = WS.Range("M2").Value 'REV
  job_Arr(8) = WS.Range("N2").Value 'DIE
  job_Arr(9) = WS.Range("O2").Value 'REEL
  job_Arr(10) = WS.Range("P2").Value 'INSP
  job_Arr(11) = WS.Range("Q2").Value 'DEVICE
  job_Arr(12) = WS.Range("B36").Value 'DATE
  job_Arr(13) = WS.Range("D36").Value 'COMMENT

  IMPORT_JOB_DATA = job_Arr

End Function

'* Write Data To PCD_ARR
'*** TESTED *** Name: Josh Kroshus Date: 10-19-21
Function IMPORT_PCD_DATA()

  Dim PCDmisFile As String: PCDmisFile = ThisWorkbook.Worksheets("SETTINGS").Range("B3").Value
  Dim WS As Worksheet: Set WS = Worksheets(PCDmisFile)
  Dim pcd_Arr(2, 50) As String
  Dim i, j, r As Integer

  For i = 1 To 2
    r = 12 'C-1 DIM A
    If i = 2 Then r = 62
      For j = 1 To 5
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 37 'C-2 DIM A
    If i = 2 Then r = 87
      For j = 26 To 30
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 13 'C-1 DIM B
    If i = 2 Then r = 63
      For j = 6 To 10
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 38 'C-2 DIM B
    If i = 2 Then r = 88
      For j = 31 To 35
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 10 'C-1 ANGLE 1
    If i = 2 Then r = 60
      For j = 11 To 15
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 35 'C-2 ANGLE 1
    If i = 2 Then r = 85
      For j = 36 To 40
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 11 'C-1 ANGLE 2
    If i = 2 Then r = 61
      For j = 16 To 20
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 36 'C-2 ANGLE 2
    If i = 2 Then r = 86
      For j = 41 To 45
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 14 'C-1 BUMP
    If i = 2 Then r = 64
      For j = 21 To 25
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
    r = 39 'C-2 BUMP
    If i = 2 Then r = 89
      For j = 46 To 50
        pcd_Arr(i, j) = WS.Cells(r, 8).Value
        r = r + 5
        Next j
  Next i
  
  IMPORT_PCD_DATA = pcd_Arr

End Function

'* 8448 WORKSHEET FUNCTIONS
'* <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>

'* Write Data To Data Collection Page (Sheet2)
'*** TESTED *** Name: Josh Kroshus Date: 10-19-21
Function IMPORT_WS_COL(pcd_Arr)

  Dim WS As Worksheet: Set WS = Worksheets("DATA COLLECTION")
  Dim i, j, r As Integer

    r = 0
    For i = 1 To 2
    If i = 2 Then r = 25
    For j = 1 To 5
      'C-1
      WS.Cells(i + 5, j + 4) = pcd_Arr(i, j)
      WS.Cells(i + 11, j + 4) = pcd_Arr(i, j + 5)
      WS.Cells(i + 17, j + 4) = pcd_Arr(i, j + 10)
      WS.Cells(i + 23, j + 4) = pcd_Arr(i, j + 15)
      WS.Cells(i + 29, j + 4) = pcd_Arr(i, j + 20)
      Next j
      Next i
    
    r = 0
    For i = 1 To 2
    If i = 2 Then r = 25
    For j = 1 To 5
      'C-2
      WS.Cells(i + 5, j + 12) = pcd_Arr(i, j + 25)
      WS.Cells(i + 11, j + 12) = pcd_Arr(i, j + 30)
      WS.Cells(i + 17, j + 12) = pcd_Arr(i, j + 35)
      WS.Cells(i + 23, j + 12) = pcd_Arr(i, j + 40)
      WS.Cells(i + 29, j + 12) = pcd_Arr(i, j + 45)
      Next j
      Next i
End Function

'* Write Data To CPK Page (Sheet3)
'*** TESTED *** Name: Josh Kroshus Date: 10-19-21
Function IMPORT_WS_CPK(job_Arr, pcd_Arr)

  Dim WS As Worksheet: Set WS = Worksheets("CPK")
  Dim i, j, lRow As Integer
  
  lRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
  For j = 1 To 5
    WS.Cells(j + lRow, 1) = job_Arr(12) 'DATE
    WS.Cells(j + lRow, 2) = job_Arr(10) 'INSP
    WS.Cells(j + lRow, 3) = job_Arr(9) 'REEL
    Next j
  i = 1
  For j = 1 To 5
    'SIDE 1 C-1
    WS.Cells(j + lRow, 4) = pcd_Arr(i, j)
    WS.Cells(j + lRow, 5) = pcd_Arr(i, j + 5)
    WS.Cells(j + lRow, 6) = pcd_Arr(i, j + 10)
    WS.Cells(j + lRow, 7) = pcd_Arr(i, j + 15)
    'SIDE 1 C-2
    WS.Cells(j + lRow, 8) = pcd_Arr(i + 1, j)
    WS.Cells(j + lRow, 9) = pcd_Arr(i + 1, j + 5)
    WS.Cells(j + lRow, 10) = pcd_Arr(i + 1, j + 10)
    WS.Cells(j + lRow, 11) = pcd_Arr(i + 1, j + 15)
    Next j
 
  lRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
  For j = 1 To 5
    WS.Cells(j + lRow, 1) = job_Arr(12) 'DATE
    WS.Cells(j + lRow, 2) = job_Arr(10) 'INSP
    WS.Cells(j + lRow, 3) = job_Arr(9) 'REEL
    Next j
  i = 1
  For j = 1 To 5
    'SIDE 2 C-1
    WS.Cells(j + lRow, 4) = pcd_Arr(i, j + 25)
    WS.Cells(j + lRow, 5) = pcd_Arr(i, j + 30)
    WS.Cells(j + lRow, 6) = pcd_Arr(i, j + 35)
    WS.Cells(j + lRow, 7) = pcd_Arr(i, j + 40)
    'SIDE 2 C-2
    WS.Cells(j + lRow, 8) = pcd_Arr(i + 1, j + 25)
    WS.Cells(j + lRow, 9) = pcd_Arr(i + 1, j + 30)
    WS.Cells(j + lRow, 10) = pcd_Arr(i + 1, j + 35)
    WS.Cells(j + lRow, 11) = pcd_Arr(i + 1, j + 40)
    Next j

End Function

'* Write Data To Data History Page (Sheet4)
'*** TESTED *** Name: Josh Kroshus Date: 10-19-21
Function IMPORT_WS_HIS(job_Arr, pcd_Arr)

  Dim WS As Worksheet: Set WS = Worksheets("DATA HISTORY")
  Dim i, j, lRow As Integer, die As String

  lRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

  For i = 1 To 2
    For j = 1 To 9
      WS.Cells(i + lRow, j) = job_Arr(j)
      Next j

    If job_Arr(8) = "C" And i = 1 = True Then WS.Cells(i + lRow, 8) = "C-1"
    If job_Arr(8) = "C" And i = 2 = True Then WS.Cells(i + lRow, 8) = "C-2"
    If i = 1 = True Then WS.Cells(i + lRow, 10) = "SIDE 1"
    If i = 2 = True Then WS.Cells(i + lRow, 10) = "SIDE 2"

    For j = 1 To 5
      WS.Cells(i + lRow, j + 15) = pcd_Arr(i, j)
      WS.Cells(i + lRow, j + 20) = pcd_Arr(i, j + 5)
      WS.Cells(i + lRow, j + 25) = pcd_Arr(i, j + 10)
      WS.Cells(i + lRow, j + 30) = pcd_Arr(i, j + 15)
      WS.Cells(i + lRow, j + 35) = pcd_Arr(i, j + 20)
      Next j
      Next i

  lRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

  For i = 1 To 2
    For j = 1 To 9
      WS.Cells(i + lRow, j) = job_Arr(j)
      Next j

    If job_Arr(8) = "C" And i = 1 = True Then WS.Cells(i + lRow, 8) = "C-1"
    If job_Arr(8) = "C" And i = 2 = True Then WS.Cells(i + lRow, 8) = "C-2"
    If i = 1 = True Then WS.Cells(i + lRow, 10) = "SIDE 1"
    If i = 2 = True Then WS.Cells(i + lRow, 10) = "SIDE 2"

    For j = 1 To 5
      WS.Cells(i + lRow, j + 15) = pcd_Arr(i, j + 25)
      WS.Cells(i + lRow, j + 20) = pcd_Arr(i, j + 30)
      WS.Cells(i + lRow, j + 25) = pcd_Arr(i, j + 35)
      WS.Cells(i + lRow, j + 30) = pcd_Arr(i, j + 40)
      WS.Cells(i + lRow, j + 35) = pcd_Arr(i, j + 45)
      Next j
      Next i

End Function

'* 8448 WORKBOOK FUNCTIONS
'* <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><> <><><><><>

'* Workbook Close and Save Function
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21
Sub WB_CLOSE()

  ThisWorkbook.Save
  If Application.Workbooks.Count = 1 Then
      Application.Quit
    Else
      ActiveWorkbook.Close
    End If

End Sub

'* Worksheet Clear Function
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21
Function WS_CLEAR(WS_NAME)

  Dim lRow, lCol As Integer
  Dim WS As Worksheet: Set WS = Sheets(WS_NAME)

  lRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
  lCol = WS.Cells(6, Columns.Count).End(xlToLeft).Column
  Range(WS.Cells(6, 1), WS.Cells(lRow, lCol)).ClearContents

End Function

'* Worksheet Delete Function
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21
Function WS_DELETE(WS_NAME)

  Application.DisplayAlerts = False
  Sheets(WS_NAME).Delete
  Application.DisplayAlerts = True

End Function

'* Worksheet Clear Data Collection (8448, 8449, 8757)
'*** TESTED *** Author: Josh Kroshus Date: 10-19-21
Function WS_CLEAR_COL()

  Dim WS As Worksheet: Set WS = Sheets("DATA COLLECTION")

  WS.[E6:I7].ClearContents
  WS.[E12:I13].ClearContents
  WS.[E18:I19].ClearContents
  WS.[E24:I25].ClearContents
  WS.[E30:I31].ClearContents

  WS.[M6:Q7].ClearContents
  WS.[M12:Q13].ClearContents
  WS.[M18:Q19].ClearContents
  WS.[M24:Q25].ClearContents
  WS.[M30:Q31].ClearContents

End Function