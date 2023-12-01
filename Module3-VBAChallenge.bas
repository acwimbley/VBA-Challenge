Attribute VB_Name = "Module3"
Sub Worksonworkbook()

'Change to work on all worksheets in workbook

For Each ws In Worksheets
Dim WorkSheetName As String

'Finding the last row of the worksheet

'Isolating the worksheet to test if script will work

'Set ws = ThisWorkbook.Sheets("A")

'Finding the last row of the worksheet

Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


Dim i As Long
Dim tickername As String
Dim totalVolume As Double
totalVolume = 0
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentChange As Double

'adding portion for collecting summary data
Dim summaryrow As Long
summaryrow = 2
Dim summaryticker As String


For i = 2 To lastRow
 
 If ws.Cells(i, 1).Value <> tickername Then
    tickername = ws.Cells(i, 1).Value

'to write to the summary sections
If i > 2 Then
 ws.Cells(summaryrow, 10).Value = summaryticker
 ws.Cells(summaryrow, 11).Value = openprice
 ws.Cells(summaryrow, 12).Value = closeprice
 ws.Cells(summaryrow, 13).Value = yearlychange
 ws.Cells(summaryrow, 14).Value = percentChange
 ws.Cells(summaryrow, 15).Value = totalVolume
 
  summaryrow = summaryrow + 1
  
End If
    
' resetting for new stock summary information
 totalVolume = 0
 openprice = ws.Cells(i, 3)
 closeprice = 0
 yearlychange = 0
 percentChange = 0
 summaryticker = tickername
 
End If

 totalVolume = totalVolume + ws.Cells(i, 7).Value
 
 closeprice = ws.Cells(i, 6).Value
 
yearlychange = closeprice - openprice

    If openprice <> 0 Then
            percentChange = yearlychange / openprice * 100
    Else
            percentChange = 0
         
    End If
   
Next i

Next ws


End Sub
