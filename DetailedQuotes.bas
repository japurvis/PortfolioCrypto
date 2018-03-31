Attribute VB_Name = "DetailedQuotes"
Private sheet As Worksheet

Sub UpdateDetailedQuotesSheet()
   
    Call DisableApplication
    Call UpdateDetailedQuotes
    Call EnableApplication
    
End Sub

Sub UpdateDetailedQuotes()
        
    Set sheet = Sheets("DetailedQuotes")
    sheet.Activate
    Application.StatusBar = "Updating Detailed Quotes"
        
    Dim row As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    For row = headerRow + 1 To lastRow
        If sheet.Cells(row, 3) <> sheet.Cells(row, 4) Then
            Call RunQuery(sheet.Cells(row, 2), sheet.Cells(row, 3), row)
        End If
    Next row
    
    Dim qt As QueryTable
    For Each qt In sheet.QueryTables
         qt.Delete
    Next qt
    
    Call PostFormat
    
    Application.StatusBar = ""
    
End Sub

Private Function RunQuery(coinName As String, quoteDate As Date, coinRow As Integer)

    Application.StatusBar = "Updating Detailed Quotes - " & sheet.Cells(coinRow, 1)
    
    Dim quoteString As String
    quoteString = Format(quoteDate, "yyyymmdd")
    
    Dim url As String
    'https://coinmarketcap.com/currencies/ethereum/historical-data/?start=20180115&end=20180121
    url = "https://coinmarketcap.com/currencies/" & coinName & "/historical-data/?start=" & quoteString & "&end=" & quoteString & ""
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    
    sheet.Rows(coinRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    With sheet.QueryTables.Add(Connection:="URL;" & url, Destination:=sheet.Cells(coinRow, 4))
        .RefreshStyle = xlOverwriteCells
        .Refresh BackgroundQuery:=False
        .SaveData = False
    End With
    
    sheet.Rows(coinRow).Delete Shift:=xlUp

End Function

Private Sub PostFormat()
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Range(sheet.Cells(headerRow + 1, 5), sheet.Cells(lastRow, lastColumn)).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    sheet.Cells(1, 1).Select
    
End Sub






