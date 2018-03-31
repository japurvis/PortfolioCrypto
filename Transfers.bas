Attribute VB_Name = "Transfers"
Private sheet As Worksheet

Sub UpdateTransfersSheet()
    
    Call DisableApplication
    Call UpdateTransfers
    Call EnableApplication
    
End Sub

Sub UpdateTransfers()

    Application.StatusBar = "Updating Transfers"
    
    Call PreFormat
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBittrex").Value) = 1 Then
        Call ApiBittrex.ParseTransfers(Sheets("Transfers"), PrivateApiBittrex("account/getdeposithistory"), "", "Bittrex")
        Call ApiBittrex.ParseTransfers(Sheets("Transfers"), PrivateApiBittrex("account/getwithdrawalhistory"), "Bittrex", "")
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBinance").Value) = 1 Then
        Call ApiBinance.ParseTransfers(Sheets("Transfers"), ApiBinance.PrivateApiBinance("GET", "depositHistory.html"), "", "Binance")
        Call ApiBinance.ParseTransfers(Sheets("Transfers"), ApiBinance.PrivateApiBinance("GET", "withdrawalHistory.html"), "Binance", "")
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataGDAX").Value) = 1 Then
        Set sheet = Sheets("Transfers")
        sheet.Activate
        
        Dim coin As String
        Dim accountId As String
        Dim row As Integer
        Dim lastRow As Integer
        lastRow = Sheets("Balances").Cells(Sheets("Balances").UsedRange.Rows.Count + 1, 1).End(xlUp).row
        
        For row = 3 To lastRow
            If Sheets("Balances").Cells(row, 2) = "GDAX" Then
                accountId = Sheets("Balances").Cells(row, 7)
                coin = Sheets("Balances").Cells(row, 3)
                Call ApiGDAX.ParseTransfers(Sheets("Transfers"), ApiGDAX.PrivateApiGDAX("GET", "/accounts/" & accountId & "/ledger"), coin)
            End If
        Next row
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataCoinbase").Value) = 1 Then
        Set sheet = Sheets("Transfers")
        sheet.Activate
        lastRow = Sheets("Balances").Cells(Sheets("Balances").UsedRange.Rows.Count + 1, 1).End(xlUp).row
        
        For row = 3 To lastRow
            If Sheets("Balances").Cells(row, 2) = "Coinbase" Then
                accountId = Sheets("Balances").Cells(row, 7)
                coin = Sheets("Balances").Cells(row, 3)
                Call ApiCoinbase.ParseTransfers(Sheets("Transfers"), ApiCoinbase.PrivateApiCoinbase("GET", "/accounts/" & accountId & "/transactions", "?&limit=100"), coin)
                'Call ApiCoinbase.ParseTransfers(Sheets("Transfers"), ApiCoinbase.PrivateApiCoinbase("GET", "/accounts/" & accountId & "/deposits", "?&limit=100"), coin)
                'Call ApiCoinbase.ParseTransfers(Sheets("Transfers"), ApiCoinbase.PrivateApiCoinbase("GET", "/accounts/" & accountId & "/withdrawals", "?&limit=100"), coin)
            End If
        Next row
    End If
    
    Call PostFormat

    Application.StatusBar = ""
    
End Sub

Sub AddTransfer(fromAcct As String, toAcct As String, coin As String, units As Double, fee As Double, fromDate As Date, toDate As Date)

    Set sheet = Sheets("Transfers")
    Dim found As Boolean
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim row As Integer
    Dim col As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 4).End(xlUp).row + 1
    col = 1
    
    For row = headerRow + 1 To lastRow
        If sheet.Cells(row, 3) = coin And Abs(sheet.Cells(row, 4)) = Abs(units) Then
            If (fromAcct <> "" And sheet.Cells(row, 1) = "" And (Abs(sheet.Cells(row, 6) - toDate)) < (60 / 86400)) Then
                found = True
                Exit For
            ElseIf (toAcct <> "" And sheet.Cells(row, 2) = "" And (Abs(sheet.Cells(row, 7) - fromDate)) < (60 / 86400)) Then
                found = True
                Exit For
            End If
        End If
    Next row
    
    If found = False Then
        For row = headerRow + 1 To lastRow
            If sheet.Cells(row, 3) = coin And Abs(sheet.Cells(row, 4)) = Abs(units) Then
                If (sheet.Cells(row, 1) = fromAcct And (Abs(sheet.Cells(row, 6) - fromDate)) < (60 / 86400)) Then
                    found = True
                    If sheet.Cells(row, 6) > 0 And sheet.Cells(row, 7) > 0 Then
                        Exit Sub
                    End If
                    Exit For
                ElseIf (sheet.Cells(row, 2) = toAcct And (Abs(sheet.Cells(row, 7) - toDate)) < (60 / 86400)) Then
                    found = True
                    If sheet.Cells(row, 6) > 0 And sheet.Cells(row, 7) > 0 Then
                        Exit Sub
                    End If
                    Exit For
                End If
            End If
        Next row
    End If
    
    If found = False Then
        row = lastRow
        sheet.Rows(row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    ElseIf found = True And sheet.Cells(row, 6) > 0 And sheet.Cells(row, 7) > 0 Then
        Exit Sub
    End If
    
    If sheet.Cells(row, col) = "" Then
        sheet.Cells(row, col) = fromAcct
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" Then
        sheet.Cells(row, col) = toAcct
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" Then
        sheet.Cells(row, col) = coin
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" Then
        sheet.Cells(row, col) = units
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" Or sheet.Cells(row, col) = 0 Then
        sheet.Cells(row, col) = fee
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" And fromDate > 0 Then
        sheet.Cells(row, col) = fromDate
    End If
    col = col + 1
    If sheet.Cells(row, col) = "" And toDate > 0 Then
        sheet.Cells(row, col) = toDate
    End If
            
End Sub

Private Sub PreFormat()

    Set sheet = Sheets("Transfers")
    sheet.Activate
    
End Sub

Private Sub PostFormat()
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 4).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    For row = headerRow + 1 To lastRow
        If sheet.Cells(row, 6) = "" Then
            sheet.Cells(row, 6) = sheet.Cells(row, 7)
        ElseIf sheet.Cells(row, 7) = "" Then
            sheet.Cells(row, 7) = sheet.Cells(row, 6)
        End If
    Next row
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 7), sheet.Cells(lastRow, 7)), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With sheet.Sort
        .SetRange Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn))
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

Sub CreateInternalTrade()
    
    'basically just need to create a buy order
    Dim transferId As String
    Dim transferFromExchange As String
    Dim transferToExchange As String
    Dim transferCurrency As String
    Dim transferUnits As Double
    Dim transferFee As Double
    Dim transferDate As Date
    Dim rate As String
    Dim row As Integer
    row = ActiveCell.row
    
    Set sheet = Sheets("Transfers")
    transferId = "*"
    transferFromExchange = sheet.Cells(row, 1)
    transferToExchange = sheet.Cells(row, 2)
    transferCurrency = sheet.Cells(row, 3)
    transferUnits = sheet.Cells(row, 4)
    transferFee = sheet.Cells(row, 5)
    transferDate = sheet.Cells(row, 7)
    
    'since anything can be transferred, need a way to look up the Open price
    rate = "=IFERROR(IF(RC4=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(RC4=""BTC"",2,IF(RC4=""ETH"",3,IF(RC4=""USDT"",4,IF(RC4=""BNB"",5,0)))),TRUE)),"""")"
    
    Call Trades.AddTrade(3, transferId, transferToExchange, "USD", transferCurrency, CStr(transferDate), CStr(transferDate), "BUY", transferUnits + transferFee, rate, "0", "0")
    
End Sub

Sub CreateExternalTrade()
    
    'basically just need to create a sell order
    Dim transferId As String
    Dim transferFromExchange As String
    Dim transferToExchange As String
    Dim transferCurrency As String
    Dim transferUnits As Double
    Dim transferFee As Double
    Dim transferDate As Date
    Dim rate As String
    Dim row As Integer
    row = ActiveCell.row
    
    Set sheet = Sheets("Transfers")
    transferId = "*"
    transferFromExchange = sheet.Cells(row, 1)
    transferToExchange = sheet.Cells(row, 2)
    transferCurrency = sheet.Cells(row, 3)
    transferUnits = sheet.Cells(row, 4)
    transferFee = sheet.Cells(row, 5)
    transferDate = sheet.Cells(row, 7)
    
    'since anything can be transferred, need a way to look up the Open price
    rate = "=IFERROR(IF(RC4=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(RC4=""BTC"",2,IF(RC4=""ETH"",3,IF(RC4=""USDT"",4,IF(RC4=""BNB"",5,0)))),TRUE)),"""")"
    
    Call Trades.AddTrade(3, transferId, transferFromExchange, "USD", transferCurrency, CStr(transferDate), CStr(transferDate), "SELL", transferUnits + transferFee, rate, "0", "0")
    
End Sub
Sub CreateIntraTrade()
    
    'need to match up all of the buy lots and flip them to the new exchange/location
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim id As String
    Dim exchange As String
    Dim baseCurrency As String
    Dim marketCurrency As String
    Dim openedDate As Date
    Dim closedDate As Date
    Dim orderType As String
    Dim units As Double
    Dim rate As String
    Dim commission As Double
    Dim additionalFees As Double
    Dim price As Double
    Dim proceeds As Double
    Dim costBasis As Double
    Dim cbpu As Double
    Dim row As Integer
    row = ActiveCell.row
    
    Set sheet = Sheets("Transfers")
    
    Dim transferId As String
    Dim transferFromExchange As String
    Dim transferToExchange As String
    Dim transferCurrency As String
    Dim transferUnits As Double
    Dim transferFee As Double
    Dim transferDate As Date
    
    transferId = "*"
    transferFromExchange = sheet.Cells(row, 1)
    transferToExchange = sheet.Cells(row, 2)
    transferCurrency = sheet.Cells(row, 3)
    transferUnits = sheet.Cells(row, 4)
    transferFee = sheet.Cells(row, 5)
    transferDate = sheet.Cells(row, 7)
    
    Set sheet = Sheets("Trades")
    sheet.Activate
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range("F:F"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With sheet.Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 2).End(xlUp).row
    
    'SELL
    If transferFee > 0 Then
        id = transferId
        exchange = transferFromExchange
        baseCurrency = "USD"
        marketCurrency = transferCurrency
        openedDate = transferDate
        closedDate = transferDate
        orderType = "SELL"
        units = transferFee
        rate = "=IFERROR(IF(RC4=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(RC4=""BTC"",2,IF(RC4=""ETH"",3,IF(RC4=""USDT"",4,IF(RC4=""BNB"",5,0)))),TRUE)),"""")"
        commission = "0"
        additionalFees = "0"
        Call Trades.AddTrade(headerRow + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
        lastRow = lastRow + 1
    End If
    
    Dim skip As Boolean
    skip = False
    
    For row = headerRow + 1 To lastRow
        If sheet.Cells(row, 6) > transferDate Or sheet.Cells(row, 2) <> transferFromExchange Then
            skip = True
        ElseIf (sheet.Cells(row, 3) <> transferCurrency And sheet.Cells(row, 4) <> transferCurrency) Then
            skip = True
        ElseIf (sheet.Cells(row, 7) = "BUY" And sheet.Cells(row, 3) = transferCurrency) Or (sheet.Cells(row, 7) = "SELL" And sheet.Cells(row, 4) = transferCurrency) Then
            skip = True
        ElseIf sheet.Cells(row, 7) = "BUY" And sheet.Cells(row, 16) = 0 Then
            skip = True
        ElseIf sheet.Cells(row, 7) = "SELL" And sheet.Cells(row, 18) = 0 Then
            skip = True
        End If
        
        If skip = False Then
            If sheet.Cells(row, 1) = sheet.Cells(row + 1, 1) Then
                If sheet.Cells(row + 1, 16) <> "" And sheet.Cells(row + 1, 16) = transferUnits Or _
                   sheet.Cells(row + 1, 18) <> "" And sheet.Cells(row + 1, 18) = transferUnits Then
                End If
            End If
            
            If sheet.Cells(row, 16) <> "" And sheet.Cells(row, 16) <= transferUnits Then
                If sheet.Cells(row, 3) = "USD" Then
                    sheet.Cells(row, 2) = transferToExchange
                    transferUnits = transferUnits - sheet.Cells(row, 16)
                ElseIf sheet.Cells(row, 16) <> sheet.Cells(row, 8) Then
                    id = sheet.Cells(row, 1)
                    exchange = sheet.Cells(row, 2)
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 3)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "SELL"
                    units = sheet.Cells(row, 13)
                    rate = sheet.Cells(row, 14) / sheet.Cells(row, 13)
                    commission = "0"
                    additionalFees = "0"
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
             
                    id = sheet.Cells(row, 1)
                    exchange = transferToExchange
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = sheet.Cells(row, 16)
                    rate = sheet.Cells(row, 9)
                    commission = Round(sheet.Cells(row, 16) / sheet.Cells(row, 8) * sheet.Cells(row, 10), 8)
                    additionalFees = Round(sheet.Cells(row, 16) / sheet.Cells(row, 8) * sheet.Cells(row, 11), 8)
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))

                    id = sheet.Cells(row, 1)
                    exchange = sheet.Cells(row, 2)
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = sheet.Cells(row, 8) - sheet.Cells(row, 16)
                    rate = sheet.Cells(row, 9)
                    commission = sheet.Cells(row, 10) - sheet.Cells(row + 1, 10)
                    additionalFees = sheet.Cells(row, 11).Text - sheet.Cells(row + 1, 11)
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                    
                    transferUnits = transferUnits - sheet.Cells(row, 16)
                    sheet.Range("A" & row).EntireRow.Delete
                Else
                    id = sheet.Cells(row, 1)
                    exchange = sheet.Cells(row, 2)
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 3)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "SELL"
                    units = sheet.Cells(row, 13)
                    rate = sheet.Cells(row, 14) / sheet.Cells(row, 13)
                    commission = "0"
                    additionalFees = "0"
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                    
                    id = sheet.Cells(row, 1)
                    exchange = transferToExchange
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = sheet.Cells(row, 16)
                    rate = sheet.Cells(row, 15)
                    commission = sheet.Cells(row, 10)
                    additionalFees = sheet.Cells(row, 11).Text
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                                                                               
                    transferUnits = transferUnits - sheet.Cells(row, 16)
                    sheet.Range("A" & row).EntireRow.Delete
                End If
            ElseIf sheet.Cells(row, 16) <> "" And sheet.Cells(row, 16) > transferUnits Then
                If sheet.Cells(row, 3) = "USD" Then
                
                    id = sheet.Cells(row, 1)
                    exchange = transferToExchange
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = transferUnits
                    rate = sheet.Cells(row, 9)
                    commission = transferUnits / sheet.Cells(row, 8) * sheet.Cells(row, 10)
                    additionalFees = sheet.Cells(row, 11).Text
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                                    
                    id = sheet.Cells(row, 1)
                    exchange = transferFromExchange
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = Round(sheet.Cells(row, 8) - transferUnits, 8)
                    rate = sheet.Cells(row, 9)
                    commission = units / sheet.Cells(row, 8) * sheet.Cells(row, 10)
                    additionalFees = sheet.Cells(row, 11).Text
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                    sheet.Cells(row + 1, 16) = sheet.Cells(row, 16) - sheet.Cells(row + 2, 16)
                    
                    transferUnits = 0
                    sheet.Range("A" & row).EntireRow.Delete
                Else
                    id = sheet.Cells(row, 1)
                    exchange = sheet.Cells(row, 2)
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 3)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "SELL"
                    units = sheet.Cells(row, 13)
                    rate = sheet.Cells(row, 14) / sheet.Cells(row, 13)
                    commission = "0"
                    additionalFees = "0"
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
             
                    id = sheet.Cells(row, 1)
                    exchange = transferToExchange
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = transferUnits
                    rate = sheet.Cells(row, 9)
                    commission = transferUnits / sheet.Cells(row, 8) * sheet.Cells(row, 10)
                    additionalFees = sheet.Cells(row, 11).Text
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))

                    id = sheet.Cells(row, 1)
                    exchange = sheet.Cells(row, 2)
                    baseCurrency = "USD"
                    marketCurrency = sheet.Cells(row, 4)
                    openedDate = sheet.Cells(row, 5)
                    closedDate = sheet.Cells(row, 6)
                    orderType = "BUY"
                    units = sheet.Cells(row, 16) - transferUnits
                    rate = sheet.Cells(row, 9)
                    commission = (sheet.Cells(row, 8) - transferUnits) / sheet.Cells(row, 8) * sheet.Cells(row, 10)
                    additionalFees = sheet.Cells(row, 11).Text
                    Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                    
                    transferUnits = 0
                    sheet.Range("A" & row).EntireRow.Delete
                End If
            ElseIf sheet.Cells(row, 18) <> "" And sheet.Cells(row, 18) = transferUnits Then
                id = sheet.Cells(row, 1)
                exchange = transferToExchange
                baseCurrency = "USD"
                marketCurrency = sheet.Cells(row, 4)
                openedDate = sheet.Cells(row, 5)
                closedDate = sheet.Cells(row, 6)
                orderType = "BUY"
                units = sheet.Cells(row, 18)
                rate = sheet.Cells(row, 13) / sheet.Cells(row, 8)
                commission = "0"
                additionalFees = "0"
                Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                                
                id = sheet.Cells(row, 1)
                exchange = sheet.Cells(row, 2)
                baseCurrency = "USD"
                marketCurrency = sheet.Cells(row, 3)
                openedDate = sheet.Cells(row, 5)
                closedDate = sheet.Cells(row, 6)
                orderType = "SELL"
                units = sheet.Cells(row, 12)
                rate = sheet.Cells(row, 13) / sheet.Cells(row, 12)
                commission = sheet.Cells(row, 10)
                additionalFees = sheet.Cells(row, 11).Text
                Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                
                transferUnits = 0
                sheet.Range("A" & row).EntireRow.Delete
            ElseIf sheet.Cells(row, 18) <> "" And sheet.Cells(row, 18) > transferUnits Then
                id = sheet.Cells(row, 1)
                exchange = sheet.Cells(row, 2)
                baseCurrency = "USD"
                marketCurrency = sheet.Cells(row, 3)
                openedDate = sheet.Cells(row, 5)
                closedDate = sheet.Cells(row, 6)
                orderType = "SELL"
                units = sheet.Cells(row, 13)
                rate = sheet.Cells(row, 14) / sheet.Cells(row, 13)
                commission = "0"
                additionalFees = "0"
                Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
        
               id = sheet.Cells(row, 1)
               exchange = transferToExchange
               baseCurrency = "USD"
               marketCurrency = sheet.Cells(row, 4)
               openedDate = sheet.Cells(row, 5)
               closedDate = sheet.Cells(row, 6)
               orderType = "BUY"
               units = transferUnits
               rate = sheet.Cells(row, 9)
               commission = transferUnits / sheet.Cells(row, 8) * sheet.Cells(row, 10)
               additionalFees = sheet.Cells(row, 11).Text
               Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))

               id = sheet.Cells(row, 1)
               exchange = sheet.Cells(row, 2)
               baseCurrency = "USD"
               marketCurrency = sheet.Cells(row, 4)
               openedDate = sheet.Cells(row, 5)
               closedDate = sheet.Cells(row, 6)
               orderType = "BUY"
               units = sheet.Cells(row, 8) - transferUnits
               rate = sheet.Cells(row, 9)
               commission = (sheet.Cells(row, 8) - transferUnits) / sheet.Cells(row, 8) * sheet.Cells(row, 10)
               additionalFees = sheet.Cells(row, 11).Text
               Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                
               transferUnits = 0
               sheet.Range("A" & row).EntireRow.Delete
            ElseIf sheet.Cells(row, 18) <> "" And sheet.Cells(row, 18) < transferUnits Then
                id = sheet.Cells(row, 1)
                exchange = transferToExchange
                baseCurrency = "USD"
                marketCurrency = sheet.Cells(row, 4)
                openedDate = sheet.Cells(row, 5)
                closedDate = sheet.Cells(row, 6)
                orderType = "BUY"
                units = sheet.Cells(row, 18)
                rate = sheet.Cells(row, 13) / sheet.Cells(row, 8)
                commission = "0"
                additionalFees = "0"
                Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                                
                id = sheet.Cells(row, 1)
                exchange = sheet.Cells(row, 2)
                baseCurrency = "USD"
                marketCurrency = sheet.Cells(row, 3)
                openedDate = sheet.Cells(row, 5)
                closedDate = sheet.Cells(row, 6)
                orderType = "SELL"
                units = sheet.Cells(row, 13) * -1
                rate = sheet.Cells(row, 14) / sheet.Cells(row, 13)
                commission = sheet.Cells(row, 10)
                additionalFees = sheet.Cells(row, 11).Text
                Call Trades.AddTrade(row + 1, id, exchange, baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, CStr(units), rate, CStr(commission), CStr(additionalFees))
                                
                transferUnits = transferUnits - sheet.Cells(row, 18)
                sheet.Range("A" & row).EntireRow.Delete
            End If
        End If
        
        If transferUnits = 0 Then
            Exit For
        End If
        
        lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 2).End(xlUp).row
        skip = False
    Next row
    
End Sub




