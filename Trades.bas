Attribute VB_Name = "Trades"
Private sheet As Worksheet

Sub UpdateTradesSheet()
    
    Call DisableApplication
    Call UpdateTrades
    Call EnableApplication
    
End Sub

Function UpdateTrades() As Integer

    Application.StatusBar = "Updating Trades"
    
    Call PreFormat
    
    Dim newTradeCount As Integer
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBittrex").Value) = 1 Then
        newTradeCount = newTradeCount + ApiBittrex.ParseTrades(Sheets("Trades"), PrivateApiBittrex("account/getorderhistory"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBinance").Value) = 1 Then
        Dim baseCurrency As String
        baseCurrency = "BTC"
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCX" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BNB", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BNB" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BTC" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-ETH", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=ETH" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-FUN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=FUN" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-GAS", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=GAS" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-IOTA", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=IOTA" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-NEO", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=NEO" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-REQ", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=REQ" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-SBTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=SBTC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-TRX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=TRX" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-VEN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=VEN" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-WTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=WTC" & baseCurrency))
        
        baseCurrency = "ETH"
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCX" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BNB", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BNB" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BTC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-ETH", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=ETH" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-FUN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=FUN" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-GAS", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=GAS" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-IOTA", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=IOTA" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-NEO", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=NEO" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-REQ", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=REQ" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-SBTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=SBTC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-TRX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=TRX" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-VEN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=VEN" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-WTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=WTC" & baseCurrency))
        
        baseCurrency = "BNB"
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BCX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BCX" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BNB", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BNB" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-BTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=BTC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-ETH", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=ETH" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-FUN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=FUN" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-GAS", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=GAS" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-IOTA", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=IOTA" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-NANO", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=NANO" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-NEO", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=NEO" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-REQ", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=REQ" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-SBTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=SBTC" & baseCurrency))
        'newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-TRX", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=TRX" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-VEN", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=VEN" & baseCurrency))
        newTradeCount = newTradeCount + ApiBinance.ParseTrades(Sheets("Trades"), baseCurrency & "-WTC", ApiBinance.PrivateApiBinance("GET", "myTrades", "symbol=WTC" & baseCurrency))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataGDAX").Value) = 1 Then
        newTradeCount = newTradeCount + ApiGDAX.ParseTrades(Sheets("Trades"), ApiGDAX.PrivateApiGDAX("GET", "/fills"))
    End If
    
    Call PostFormat
    
    UpdateTrades = newTradeCount
    Application.StatusBar = ""
    
End Function

Sub ResetCapitalGains()

    Call DisableApplication
    Call CapitalGains.ResetCapitalGains
    Call EnableApplication
    
End Sub

Sub AddTrade(row As Integer, id As String, exchange As String, baseCurrency As String, marketCurrency As String, _
    openedDate As String, closedDate As String, tradeType As String, units As String, rate As String, commission As String, additionalFees As String)
    
    Dim curSheet As Worksheet
    Set curSheet = ActiveSheet
    
    Set sheet = Sheets("Trades")
    sheet.Activate
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
    sheet.Cells(row, col) = openedDate
    col = col + 1
    sheet.Cells(row, col) = closedDate
    col = col + 1
    sheet.Cells(row, col) = tradeType
    col = col + 1
    If InStr(1, units, "=") = 1 Then
        sheet.Cells(row, col).FormulaR1C1 = units
    Else
        sheet.Cells(row, col) = units
    End If
    col = col + 1
    If InStr(1, rate, "=") = 1 Then
        sheet.Cells(row, col).FormulaR1C1 = rate
    Else
        sheet.Cells(row, col) = rate
    End If
    col = col + 1
    If InStr(1, commission, "=") = 1 Then
        sheet.Cells(row, col).FormulaR1C1 = commission
    Else
        sheet.Cells(row, col) = commission
    End If
    col = col + 1
    If InStr(1, additionalFees, "=") = 1 Then
        sheet.Cells(row, col).FormulaR1C1 = additionalFees
    Else
        sheet.Cells(row, col) = additionalFees
    End If
    col = col + 1
    If exchange = "Bittrex" Then
        sheet.Cells(row, col).FormulaR1C1 = "=ROUNDUP((RC8*RC9)+RC[-1],8)"
    ElseIf exchange = "Binance" Then
        sheet.Cells(row, col).FormulaR1C1 = "=ROUNDDOWN((RC8*RC9)+RC[-1],8)"
    Else
        sheet.Cells(row, col).FormulaR1C1 = "=ROUNDUP((RC8*RC9)+RC[-1],8)"
    End If
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IFERROR(IF(RC7=""BUY"",RC[-1]+RC[-3],RC[-1]-RC[-3])*(IF(RC7=""SELL"",-1,1)),"""")"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IFERROR(RC[-1]*IF(RC3=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(RC3=""BTC"",2,IF(RC3=""ETH"",3,IF(RC3=""USDT"",4,IF(RC3=""BNB"",5,0)))),TRUE)),"""")"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IFERROR(RC[-1]/RC8,"""")"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IF(RC7=""SELL"","""",RC8)"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]*RC[-1],"""")"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IF(RC7=""BUY"","""",RC13*-1)"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = "=IF(RC[-1]<>"""",(RC14/RC13)*RC[-1],"""")"
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = ""
    col = col + 1
    sheet.Cells(row, col).FormulaR1C1 = ""
            
End Sub

Private Sub PreFormat()
    
    Set sheet = Sheets("Trades")
    sheet.Activate
    
End Sub

Private Sub PostFormat()
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    sheet.Activate
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 6), sheet.Cells(lastRow, 6)), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
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

Sub CalculateCapitalGains()

    Call DisableApplication
    Call CapitalGains.CalculateCapitalGains
    Call EnableApplication
    
End Sub

