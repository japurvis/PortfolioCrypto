Attribute VB_Name = "Import"
Private sheet As Worksheet

Sub ImportTrades()

    Call DisableApplication
    
    Dim tradesSheet As Worksheet
    Dim tradesRow As Integer
    Dim tradesLastRow As Integer
    Dim importSheet As Worksheet
    Dim importRow As Integer
    Dim importLastRow As Integer
    Dim headerRow As Integer
    Dim found As Boolean
    
    Dim id As String
    Dim quoteId As String
    Dim baseCurrency As String
    Dim marketCurrency As String
    Dim openedDate As Date
    Dim closedDate As Date
    Dim orderType As String
    Dim units As String
    Dim rate As String
    Dim commission As String
    Dim price As String
    Dim additionalFees As String
    
    Set tradesSheet = Sheets("Trades")
    tradesSheet.Activate
    tradesLastRow = tradesSheet.Cells(tradesSheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    Set importSheet = Sheets("Import")
    importSheet.Activate
    importLastRow = importSheet.Cells(importSheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    headerRow = 2
    found = False
    
    For importRow = importLastRow To headerRow + 1 Step -1
        Set sheet = importSheet
        sheet.Activate
        found = False
        
        id = sheet.Cells(importRow, 1)
        baseCurrency = Split(sheet.Cells(importRow, 2), "-")(0)
        marketCurrency = Split(sheet.Cells(importRow, 2), "-")(1)
        orderType = Split(sheet.Cells(importRow, 3), "_")(1)
        units = Round(sheet.Cells(importRow, 4), 8)
        rate = Round(sheet.Cells(importRow, 7) / sheet.Cells(importRow, 4), 8)
        commission = sheet.Cells(importRow, 6)
        additionalFees = "0"
        openedDate = sheet.Cells(importRow, 8) - 0.25
        closedDate = sheet.Cells(importRow, 9) - 0.25
        
        Set sheet = tradesSheet
        sheet.Activate
        For tradesRow = headerRow + 1 To tradesLastRow
            If id = sheet.Cells(tradesRow, 1) Then
                found = True
                Exit For
            End If
        Next tradesRow
        
        If found = False Then
            Call Trades.AddTrade(tradesLastRow + 1, id, "Bittrex", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, additionalFees)
            Call Portfolio.AddCurrency(marketCurrency, "Bittrex")
            Call Portfolio.AddCurrency(baseCurrency, "Bittrex")
            Call Portfolio.AddMostRecentTrade("Bittrex", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
            Call Dashboard.AddCurrency(marketCurrency, "Bittrex")
            Call Dashboard.AddCurrency(baseCurrency, "Bittrex")
        End If
        
        id = ""
        baseCurrency = ""
        marketCurrency = ""
        orderType = ""
        units = ""
        rate = ""
        commission = ""
        additionalFees = 0
        openedDate = 0
        closedDate = 0
        Set sheet = importSheet
        sheet.Range("A" & importRow).EntireRow.Delete
    Next importRow
    
    Call EnableApplication
End Sub


