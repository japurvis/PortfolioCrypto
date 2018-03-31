Attribute VB_Name = "ApiBinance"
'https://www.binance.com/restapipub.html#grip-content

Function PublicApiBinance(endpoint As String, Optional options As String) As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://api.binance.com/api/v1/" & endpoint & options
    http.Send
    http.WaitForResponse
    PublicApiBinance = http.ResponseText
    Set http = Nothing

End Function

Function PrivateApiBinance(Method As String, endpoint As String, Optional options As String) As String
    
    Dim resultList() As String
    Dim apikey As String
    Dim apisecret As String
    Dim signature As String
    Dim timestamp As String
    Dim serverTime As String
    Dim apiInfo As String
    
    If InStr(endpoint, ".html") > 0 Then
        apiInfo = "wapi"
    Else: apiInfo = "api"
    End If
    
    apikey = Evaluate(ActiveWorkbook.Names("ApiKeyBinance").Value)
    apisecret = Evaluate(ActiveWorkbook.Names("ApiSecretBinance").Value)
    If Trim(apikey) = "" Or Trim(apisecret) = "" Then
        Exit Function
    End If
    
    apiUrl = "https://api.binance.com/" & apiInfo & "/v3/" & endpoint
    resultList = Split(PublicApiBinance("/time", ""), ":")
    timestamp = Replace(resultList(UBound(resultList)), "}", "")
    queryString = options & "&recvWindow=5000&timestamp=" & timestamp
    signature = ComputeHash_C("SHA256", queryString, apisecret, "STRHEX")
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open UCase(Method), apiUrl + "?" + queryString & "&signature=" & signature, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "X-MBX-APIKEY", apikey
    http.Send
    http.WaitForResponse
    PrivateApiBinance = http.ResponseText
    Set http = Nothing

End Function

Sub TestBinance()
    Dim str As String
    
    str = PrivateApiBinance("GET", "allOrders", "orderId=1")
    'str = PublicApiBinance("/ping", "")
    'str = PublicApiBinance("/time", "")
    MsgBox (str)
    
End Sub

Sub CancelOrder(symbol As String, Optional orderId As String)
    
    If symbol = "BCH" Then
        symbol = "BCC"
    End If
    
    Call PrivateApiBinance("DELETE", "order", "symbol=" & symbol & "&orderId=" & orderId)
    
End Sub

Function PlaceOrder(baseCurrency As String, marketCurrency As String, quantity As Double, rate As Double) As String
    If quantity > 0 Then
        PlaceOrder = PlaceBuyOrder(market, quantity, rate)
    ElseIf quantity < 0 Then
        PlaceOrder = PlaceSellOrder(market, quantity * -1, rate)
    End If
    
End Function

Sub PlaceBuyOrder(baseCurrency As String, marketCurrency As String, orderType As String, timeInForce As String, quantity As Double, price As Double)
    
    PlaceBuyOrder = PrivateApiBinance("POST", "order/test")
    
End Sub

Sub PlaceSellOrder(baseCurrency As String, marketCurrency As String, orderType As String, timeInForce As String, quantity As Double, price As Double)
        
    PlaceSellOrder = PrivateApiBinance("POST", "order/test")
    
End Sub

Sub ParseBalances(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Balances - Binance"
        
    Dim jsonObject() As String
    Dim resultList() As String
    Dim str As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim marketCurrency As String
    Dim totalUnits As String
    Dim availableUnits As String
    Dim pendingUnits As String
    
    sheet.Activate
    headerRow = 2
    
    jsonObject = Split(jsonString, ":")
    If UBound(jsonObject) < 3 Then
        Exit Sub
    End If
       
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = 0 To UBound(jsonObject)
        
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            obj = Split(resultList(1), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
                       
            If (obj(1) <> "0.00000000") Then
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = Replace(obj(0), """", "")
                    obj(1) = Replace(obj(1), """", "")
                    If UCase(obj(0)) = "ASSET" Then
                        If UCase(obj(1)) = "BCC" Then
                            marketCurrency = "BCH"
                        Else
                            marketCurrency = obj(1)
                        End If
                    ElseIf UCase(obj(0)) = "FREE" Then
                        availableUnits = obj(1)
                    ElseIf UCase(obj(0)) = "LOCKED" Then
                        pendingUnits = obj(1)
                    End If
                Next j
                
                totalUnits = CDbl(availableUnits) + CDbl(pendingUnits)
                Call Balances.AddBalance(headerRow + 1, "Binance", marketCurrency, totalUnits, availableUnits, pendingUnits, "")
                Call Dashboard.AddCurrency(marketCurrency, "Binance")
                
                marketCurrency = ""
                totalUnits = ""
                availableUnits = ""
                pendingUnits = ""
            End If
        Next i
    End If
End Sub

Function ParseTrades(sheet As Worksheet, currencyPair As String, jsonString As String) As Integer

    Application.StatusBar = "Updating Trades - Binance - " & currencyPair
    
    Dim jsonObject() As String
    Dim resultList() As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim newTradeCount As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim commissionAsset As String
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
    
    sheet.Activate
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 11).End(xlUp).row
    
    jsonObject = Split(jsonString, ":")
    If UBound(jsonObject) < 3 Then
        Exit Function
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = UBound(jsonObject) To 0 Step -1
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            obj = Split(resultList(0), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
            
            found = False
            For row = headerRow + 1 To lastRow
                If obj(1) = sheet.Cells(row, 1) Then
                    found = True
                    Exit For
                End If
            Next row
            
            If found = False Then
                obj = Split(currencyPair, "-")
                baseCurrency = obj(0)
                If UCase(obj(1)) = "BCC" Then
                    marketCurrency = "BCH"
                Else
                    marketCurrency = obj(1)
                End If
                newTradeCount = newTradeCount + 1
                                
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = Replace(obj(0), """", "")
                    obj(1) = Replace(obj(1), """", "")
                    If (obj(1) <> "null" And obj(1) <> "NONE") Then
                        
                        If UCase(obj(0)) = "ID" Then
                            id = obj(1)
                        ElseIf UCase(obj(0)) = "TIME" Then
                            openedDate = UnixTimeToDate((obj(1) / 1000) - (6 * 60 * 60))
                            closedDate = UnixTimeToDate((obj(1) / 1000) - (6 * 60 * 60))
                        ElseIf UCase(obj(0)) = "ISBUYER" Then
                            If UCase(obj(1)) = "TRUE" Then
                                orderType = "BUY"
                            ElseIf UCase(obj(1)) = "FALSE" Then
                                orderType = "SELL"
                            End If
                        ElseIf UCase(obj(0)) = "QTY" Then
                            units = UCase(obj(1))
                        ElseIf UCase(obj(0)) = "PRICE" Then
                            rate = UCase(obj(1))
                        ElseIf UCase(obj(0)) = "COMMISSION" Then
                            commission = UCase(obj(1))
                        ElseIf UCase(obj(0)) = "COMMISSIONASSET" Then
                            commissionAsset = UCase(obj(1))
                        End If
                    End If
                Next j
                       
                If commissionAsset = baseCurrency Then
                   
                    Call Trades.AddTrade(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, "0")
                    Call Portfolio.AddCurrency(marketCurrency, "Binance")
                    Call Portfolio.AddCurrency(baseCurrency, "Binance")
                    Call Portfolio.AddMostRecentTrade("Binance", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                    Call Dashboard.AddCurrency(marketCurrency, "Binance")
                    Call Dashboard.AddCurrency(baseCurrency, "Binance")

                ElseIf commissionAsset = marketCurrency Then
                
                    Call Trades.AddTrade(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, "0", "0")
                    Call Portfolio.AddCurrency(marketCurrency, "Binance")
                    Call Portfolio.AddCurrency(baseCurrency, "Binance")
                    Call Portfolio.AddMostRecentTrade("Binance", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                    Call Dashboard.AddCurrency(marketCurrency, "Binance")
                    Call Dashboard.AddCurrency(baseCurrency, "Binance")
                                       
                    rate = "=IFERROR(" & sheet.Cells(headerRow + 1, 12) & "*IF(""" & baseCurrency & """=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(""" & baseCurrency & """=""BTC"",2,IF(""" & baseCurrency & """=""ETH"",3,IF(""" & baseCurrency & """=""USDT"",4,IF(""" & baseCurrency & """=""BNB"",5,0)))),TRUE))/" & sheet.Cells(headerRow + 1, 8) & ","""")"
                    baseCurrency = "USD"
                    marketCurrency = commissionAsset
                    orderType = "SELL"
                    units = commission
                    commission = 0
                    
                    Call Trades.AddTrade(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, "0")
                    Call Portfolio.AddCurrency(marketCurrency, "Binance")
                    Call Portfolio.AddCurrency(baseCurrency, "Binance")
                    Call Portfolio.AddMostRecentTrade("Binance", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                    Call Dashboard.AddCurrency(marketCurrency, "Binance")
                    Call Dashboard.AddCurrency(baseCurrency, "Binance")
                Else
                    Call Trades.AddTrade(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, "0", "0")
                    Call Portfolio.AddCurrency(marketCurrency, "Binance")
                    Call Portfolio.AddCurrency(baseCurrency, "Binance")
                    Call Portfolio.AddMostRecentTrade("Binance", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                    Call Dashboard.AddCurrency(marketCurrency, "Binance")
                    Call Dashboard.AddCurrency(baseCurrency, "Binance")
                    
                    baseCurrency = "USD"
                    marketCurrency = commissionAsset
                    orderType = "SELL"
                    units = commission
                    rate = "=IFERROR(IF(RC4=""USD"",1,VLOOKUP(RC6,HistoricalQuotes,IF(RC4=""BTC"",2,IF(RC4=""ETH"",3,IF(RC4=""USDT"",4,IF(RC4=""BNB"",5,0)))),TRUE)),"""")"
                    commission = 0
                    
                    Call Trades.AddTrade(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, "0")
                    Call Portfolio.AddCurrency(marketCurrency, "Binance")
                    Call Portfolio.AddCurrency(baseCurrency, "Binance")
                    Call Portfolio.AddMostRecentTrade("Binance", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                    Call Dashboard.AddCurrency(marketCurrency, "Binance")
                    Call Dashboard.AddCurrency(baseCurrency, "Binance")
                End If
                
                id = ""
                baseCurrency = ""
                marketCurrency = ""
                orderType = ""
                units = ""
                rate = ""
                commission = ""

            End If
        Next i
    End If
    
    ParseTrades = newTradeCount
    
End Function

Sub ParseOrders(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Orders - Binance"
        
    Dim jsonObject() As String
    Dim resultList() As String
    Dim str As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim id As String
    Dim baseCurrency As String
    Dim marketCurrency As String
    Dim orderType As String
    Dim units As String
    Dim limit As String
    Dim openedDate As Date
    
    sheet.Activate
    headerRow = 2

    jsonObject = Split(jsonString, ":")
    If UBound(jsonObject) < 3 Then
        Exit Sub
    End If
       
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
    
        For i = 0 To UBound(jsonObject)
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            obj = Split(resultList(0), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
                       
            For j = 0 To UBound(resultList)
                obj = Split(resultList(j), """:")
                obj(0) = Replace(obj(0), """", "")
                obj(1) = Replace(obj(1), """", "")
                If (obj(1) <> "null" And obj(1) <> "NONE") Then
                    If UCase(obj(0)) = "SYMBOL" Then
                        If Len(obj(1)) = 6 Then
                            marketCurrency = Left(obj(1), 3)
                            baseCurrency = Right(obj(1), 3)
                        ElseIf Len(obj(1)) = 8 Then
                            marketCurrency = Left(obj(1), 4)
                            baseCurrency = Right(obj(1), 4)
                        Else
                            If Right(obj(1), 4) = "USDT" Then
                                marketCurrency = Left(obj(1), 3)
                                baseCurrency = Right(obj(1), 4)
                            Else
                                marketCurrency = Left(obj(1), 4)
                                baseCurrency = Right(obj(1), 3)
                            End If
                        End If
                        
                        If marketCurrency = "BCC" Then
                            marketCurrency = "BCH"
                        End If
                    ElseIf UCase(obj(0)) = "ORDERID" Then
                        id = obj(1)
                    ElseIf UCase(obj(0)) = "SIDE" Then
                        orderType = obj(1)
                    ElseIf UCase(obj(0)) = "ORIGQTY" Then
                        units = obj(1)
                    ElseIf UCase(obj(0)) = "PRICE" Then
                        limit = obj(1)
                    ElseIf UCase(obj(0)) = "TIME" Then
                        openedDate = UnixTimeToDate((obj(1) / 1000) - (6 * 60 * 60))
                    End If
                End If
            Next j
                
            Call Orders.AddOrder(headerRow + 1, id, "Binance", baseCurrency, marketCurrency, orderType, units, limit, CStr(openedDate))
                
            id = ""
            baseCurrency = ""
            marketCurrency = ""
            orderType = ""
            units = ""
            limit = ""
        Next i
    End If
    
End Sub

Sub ParseTransfers(sheet As Worksheet, jsonString As String, fromAcct As String, toAcct As String)
       
    Application.StatusBar = "Updating Transfers - Binance"
    
    Dim jsonObject() As String
    Dim resultList() As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim found As Boolean
    Dim coin As String
    Dim units As Double
    Dim fee As Double
    Dim fromDate As Date
    Dim toDate As Date
    Dim transferDate As Date
    
    sheet.Activate
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    jsonObject = Split(jsonString, "success"":true")
    If UBound(jsonObject) < 1 Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = UBound(jsonObject) To 0 Step -1
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            'Call Helpers.WriteDataToTest(resultList)
            
            obj = Split(resultList(4), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
            
            found = False
            For row = headerRow + 1 To lastRow
                If obj(1) = sheet.Cells(row, 8) Then
                    found = True
                    Exit For
                End If
            Next row
            
            If found = False Then
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = Replace(obj(0), """", "")
                    obj(1) = Replace(obj(1), """", "")
                    If (obj(1) <> "null") Then
                        If UCase(obj(0)) = "ASSET" Then
                            If UCase(obj(1)) = "BCC" Then
                                coin = "BCH"
                            Else
                                coin = obj(1)
                            End If
                        ElseIf UCase(obj(0)) = "AMOUNT" Then
                            units = obj(1)
                        ElseIf UCase(obj(0)) = "INSERTTIME" Then
                            transferDate = UnixTimeToDate((obj(1) / 1000) - (6 * 60 * 60))
                        End If
                    End If
                Next j
                
                If UCase(fromAcct) = "BINANCE" Then
                    fromDate = transferDate
                ElseIf UCase(toAcct) = "BINANCE" Then
                    toDate = transferDate
                End If
                
                Call Transfers.AddTransfer(fromAcct, toAcct, coin, Abs(units), fee, fromDate, toDate)
                
                coin = ""
                units = 0
                fee = 0
                fromDate = 0
                toDate = 0
                transferDate = 0
            End If
        Next i
    End If
    
End Sub
