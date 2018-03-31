Attribute VB_Name = "Orders"
Private sheet As Worksheet

Sub UpdateOrdersSheet()
    
    Call DisableApplication
    Call UpdateOrders
    Call EnableApplication
    
End Sub

Sub UpdateOrders()

    Application.StatusBar = "Updating Orders"
    
    Call PreFormat
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBittrex").Value) = 1 Then
        Call ApiBittrex.ParseOrders(Sheets("Orders"), PrivateApiBittrex("market/getopenorders"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBinance").Value) = 1 Then
        Call ApiBinance.ParseOrders(Sheets("Orders"), ApiBinance.PrivateApiBinance("GET", "openOrders"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataGDAX").Value) = 1 Then
        Call ApiGDAX.ParseOrders(Sheets("Orders"), ApiGDAX.PrivateApiGDAX("GET", "/orders"))
    End If
    
    Call PostFormat

    Application.StatusBar = ""
    
End Sub

Sub AddOrder(row As Integer, id As String, exchange As String, baseCurrency As String, marketCurrency As String, _
    orderType As String, units As String, limit As String, openedDate As Date)
    
    Set sheet = Sheets("Orders")
    Dim col As Integer
    col = 1
    
    sheet.Rows(row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    sheet.Cells(row, col) = id
    col = col + 1
    sheet.Cells(row, col) = exchange
    col = col + 1
    sheet.Cells(row, col) = baseCurrency
    col = col + 1
    sheet.Cells(row, col) = marketCurrency
    col = col + 1
    sheet.Cells(row, col) = orderType
    col = col + 1
    sheet.Cells(row, col) = units
    col = col + 1
    sheet.Cells(row, col) = limit
    col = col + 1
    sheet.Cells(row, col) = openedDate
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IFERROR((RC7-VLOOKUP(RC2&""-""&RC3&""-""&RC4,Quotes,7,FALSE))/VLOOKUP(RC2&""-""&RC3&""-""&RC4,Quotes,7,FALSE),"""")"
            
End Sub

Private Sub PreFormat()
    
    Set sheet = Sheets("Orders")
    sheet.Activate
        
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.UsedRange.Rows.Count
    lastColumn = sheet.UsedRange.Columns.Count
    
    If lastRow > headerRow Then
        Rows("" & headerRow + 1 & ":" & lastRow & "").Delete Shift:=xlUp
    End If
    
End Sub

Private Sub PostFormat()
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 2), sheet.Cells(lastRow, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 4), sheet.Cells(lastRow, 4)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 3), sheet.Cells(lastRow, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 7), sheet.Cells(lastRow, 7)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sheet.Sort
        .SetRange Range(sheet.Cells(headerRow, 1), sheet.Cells(sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn)).Font.Bold = True
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn)).EntireColumn.AutoFit
    sheet.Cells(1, 1).Select
End Sub

Sub NewOrder()
    
End Sub

Sub CancelOrders()
    Dim CancelOrder As Boolean
    Dim message As String
    Dim orderId As String
    Set sheet = Sheets("Orders")
    
    Dim exchange As String
    Dim currencyPair As String
    Dim quantity As Double
    Dim price As Double
    Dim orderType As String
    
    Call DisableApplication
    
    Dim r As Range
    For Each r In Selection.Rows
        
        exchange = sheet.Cells(r.row, 2)
        currencyPair = sheet.Cells(r.row, 3) & "-" & sheet.Cells(r.row, 4)
        orderType = sheet.Cells(r.row, 5)
        quantity = sheet.Cells(r.row, 6)
        price = sheet.Cells(r.row, 7)
        
        message = "Cancel " & orderType & " order on " & exchange & " for " & quantity & " units of " & currencyPair & " @ " & price & " ?"
        If MsgBox(message, vbYesNo) = vbYes Then
            CancelOrder = True
        End If
        
        If CancelOrder = True Then
            orderId = sheet.Cells(r.row, 1)
            If exchange = "Bittrex" Then
                Call ApiBittrex.CancelOrder(orderId)
            ElseIf exchange = "Binance" Then
                Call ApiBinance.CancelOrder(sheet.Cells(r.row, 4) & sheet.Cells(r.row, 3), orderId)
            End If
        End If
        
        CancelOrder = False
    Next r
    
    Call UpdateOrders
    'Call UpdateOrders
    Call EnableApplication
    
End Sub




