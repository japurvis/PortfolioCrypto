Attribute VB_Name = "HistoricalQuotes"
Private sheet As Worksheet

Sub UpdateHistoricalQuotes()
    
    Application.StatusBar = "Updating Historical Quotes"
    
    Set sheet = Sheets("HistoricalQuotes")
    sheet.Activate
    
    Dim now As Date
    Dim d As Date
    now = Date
    d = sheet.Cells(sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row, 1)
    If Not IsDate(d) Or (Year(d) <> Year(now) Or Month(d) <> Month(now) Or Day(d) <> Day(now)) Then
        ActiveWorkbook.Connections("Query - HistoricalQuotes").Refresh
    End If
    
    Application.StatusBar = ""
    
End Sub

