'! Author: Josh Kroshus  Date: 10-19-21 FOR DEBUGGING ONLY
Sub CF_DELETE()
    
    'Worksheets("CPK").Cells.FormatConditions.Delete
    'Worksheets("DATA HISTORY").Range("K6:AN1000").FormatConditions.Delete

End Sub

'*** TESTED *** Author: Josh Kroshus  Date: 06-23-21
Sub CF_COL_xlBlanks()

  Dim JOB As Range: Set JOB = Worksheets("DATA COLLECTION").Range("D2,F2,H2,J2,L2:Q2")

  JOB.FormatConditions.Delete
  JOB.FormatConditions.Add Type:=xlBlanksCondition
  JOB.FormatConditions(1).Interior.Color = RGB(255, 163, 163)

End Sub

'*** TESTED *** Author: Josh Kroshus  Date: 06-23-21
Sub CF_COL_xlBetween()

    Dim NVS231 As Range: Set NVS231 = Worksheets("DATA COLLECTION").Range("E6:I7,M6:Q7")
    Dim NVS168 As Range: Set NVS168 = Worksheets("DATA COLLECTION").Range("E12:I13,M12:Q13")
    Dim ANGLE1 As Range: Set ANGLE1 = Worksheets("DATA COLLECTION").Range("E18:I19,M18:Q19")
    Dim ANGLE2 As Range: Set ANGLE2 = Worksheets("DATA COLLECTION").Range("E24:I25,M24:Q25")
    Dim BUMP_H As Range: Set BUMP_H = Worksheets("DATA COLLECTION").Range("E30:I31,M30:Q31")

    With NVS231
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$C$8", Formula2:="=$C$9"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$C$8", Formula2:="=$C$9"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With NVS168
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$C$14", Formula2:="=$C$15"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$C$14", Formula2:="=$C$15"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE1
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$C$20", Formula2:="=$C$21"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$C$20", Formula2:="=$C$21"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE2
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$C$26", Formula2:="=$C$27"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$C$26", Formula2:="=$C$27"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With BUMP_H
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$C$32", Formula2:="=$C$33"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$C$32", Formula2:="=$C$33"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With
    
End Sub

'*** TESTED *** Author: Josh Kroshus  Date: 06-23-21
Sub CF_CPK_xlBetween()

    Dim NVS231 As Range: Set NVS231 = Worksheets("CPK").Range("D6:D1000, H6:H1000")
    Dim NVS168 As Range: Set NVS168 = Worksheets("CPK").Range("E6:E1000, I6:I1000")
    Dim ANGLE1 As Range: Set ANGLE1 = Worksheets("CPK").Range("F6:F1000, J6:J1000")
    Dim ANGLE2 As Range: Set ANGLE2 = Worksheets("CPK").Range("G6:G1000, K6:K1000")

    With NVS231
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$D$3", Formula2:="=$D$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$D$3", Formula2:="=$D$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With NVS168
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$E$3", Formula2:="=$E$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$E$3", Formula2:="=$E$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE1
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$F$3", Formula2:="=$F$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$F$3", Formula2:="=$F$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE2
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$G$3", Formula2:="=$G$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$G$3", Formula2:="=$G$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With
    
End Sub

'*** TESTED *** Author: Josh Kroshus  Date: 06-23-21
Sub CF_HIS_xlBetween()

    Dim NVS231 As Range: Set NVS231 = Worksheets("DATA HISTORY").Range("P6:T1000")
    Dim NVS168 As Range: Set NVS168 = Worksheets("DATA HISTORY").Range("U6:Y1000")
    Dim ANGLE1 As Range: Set ANGLE1 = Worksheets("DATA HISTORY").Range("Z6:AD1000")
    Dim ANGLE2 As Range: Set ANGLE2 = Worksheets("DATA HISTORY").Range("AE6:AI1000")
    Dim BUMP_H As Range: Set BUMP_H = Worksheets("DATA HISTORY").Range("AJ6:AN1000")

    With NVS231
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$P$3", Formula2:="=$P$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$P$3", Formula2:="=$P$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With NVS168
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$U$3", Formula2:="=$U$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$U$3", Formula2:="=$U$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE1
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$Z$3", Formula2:="=$Z$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$Z$3", Formula2:="=$Z$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With ANGLE2
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$AE$3", Formula2:="=$AE$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$AE$3", Formula2:="=$AE$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With

    With BUMP_H
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=$AJ$3", Formula2:="=$AJ$4"
        .FormatConditions(1).Font.Color = RGB(0, 180, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlNotBetween, Formula1:="=$AJ$3", Formula2:="=$AJ$4"
        .FormatConditions(2).Font.Color = RGB(255, 0, 0)
    End With
    
End Sub