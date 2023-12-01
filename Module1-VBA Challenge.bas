Attribute VB_Name = "Module1"

'Combining all worksheets into one summary sheet

Sub SummarySheet()

    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Stock_Data"
    'move created sheet to be first sheet
    Sheets("Stock_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Stock_Data")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRowticker = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        ' Copy the contents of each stock sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowticker - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowticker + 1)).Value

    Next ws

    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:G").AutoFit
    
    
    

End Sub
