Attribute VB_Name = "Module4"
Sub ConditionalFormatting()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stock_Data")

    ' Assuming your yearly change data starts from column 13 (adjust as needed)
    Dim yearlyColumn As Range
    Set yearlyColumn = ws.Range("M2:M" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Clear existing conditional formatting
    yearlyColumn.FormatConditions.Delete

    ' Set the rule for positive values (green)
    With yearlyColumn.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green
    End With

    ' Set the rule for negative values (red)
    With yearlyColumn.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red
    End With
    
        ' Assuming your percent change data starts from column 14 (adjust as needed)
    Dim percentColumn As Range
    Set percentColumn = ws.Range("N2:N" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Clear existing conditional formatting
    percentColumn.FormatConditions.Delete
    
      ' Set the rule for positive values (green)
    With percentColumn.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green
    End With

    ' Set the rule for negative values (red)
    With percentColumn.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red
    End With

End Sub

