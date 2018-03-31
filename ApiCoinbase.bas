Attribute VB_Name = "ApiCoinbase"

Function PublicApiCoinbase(endpoint As String, Optional MethodOptions As String) As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://api.coinbase.com/v2" & endpoint & MethodOptions
    http.SetRequestHeader "CB-VERSION", "2017-08-07"
    http.Send
    http.WaitForResponse
    PublicApiCoinbase = http.ResponseText
    Set http = Nothing

End Function

Function PrivateApiCoinbase(Method As String, endpoint As String, Optional options As String) As String
    
    Dim resultList() As String
    Dim url As String
    Dim version As String
    Dim key As String
    Dim secret As String
    Dim timestamp As String
    Dim requestPath As String
    Dim body As String
    Dim hmac As String
    Dim postdata As String
    Dim signature As String
    
    version = "/v2"
    key = Evaluate(ActiveWorkbook.Names("ApiKeyCoinbase").Value)
    secret = Evaluate(ActiveWorkbook.Names("ApiSecretCoinbase").Value)
    If Trim(key) = "" Or Trim(secret) = "" Then
        Exit Function
    End If
    
    resultList = Split(PublicApiCoinbase("/time", ""), ":")
    timestamp = Replace(resultList(UBound(resultList)), "}", "")
    
    url = "https://api.coinbase.com" & version & endpoint & options
    postdata = timestamp & UCase(Method) & version & endpoint & options
    signature = ComputeHash_C("SHA256", postdata, secret, "STRHEX")
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open UCase(Method), url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "CB-ACCESS-KEY", key
    http.SetRequestHeader "CB-ACCESS-SIGN", signature
    http.SetRequestHeader "CB-ACCESS-TIMESTAMP", timestamp
    http.SetRequestHeader "CB-VERSION", "2017-08-07"
    http.Send
    http.WaitForResponse
    PrivateApiCoinbase = http.ResponseText
    Set http = Nothing

End Function

Sub TestCoinbase()
    Dim str As String
    'str = ApiCoinbase.PrivateApiCoinbase("GET", "/accounts", "?&limit=100")
    'MsgBox (str)
    
    Call ApiCoinbase.ParseBalances(Sheets("Balances"), ApiCoinbase.PrivateApiCoinbase("GET", "/accounts", "?&limit=100"))
    
End Sub

Sub ParseBalances(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Balances - Coinbase"
    
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
    Dim accountId As String
        
    sheet.Activate
    headerRow = 2
    
    If jsonString = "The service is unavailable." Then
        Exit Sub
    End If
        
    jsonObject = Split(jsonString, "data"":")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = 0 To UBound(jsonObject)
        
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            jsonObject(i) = Replace(jsonObject(i), "[", "")
            resultList = Split(jsonObject(i), ",")
                        
            For j = 0 To UBound(resultList)
                If InStr(resultList(j), """:") > 1 Then
                    obj = Split(resultList(j), """:")
                    obj(0) = UCase(Replace(obj(0), """", ""))
                    obj(1) = Replace(obj(1), """", "")
                    If obj(0) = "ID" Then
                        accountId = obj(1)
                    ElseIf obj(0) = "CURRENCY" Then
                        marketCurrency = Replace(obj(UBound(obj)), """", "")
                    ElseIf obj(0) = "BALANCE" Then
                        totalUnits = Replace(obj(UBound(obj)), """", "")
                    End If
                End If
            Next j
            
            availableUnits = 0
            pendingUnits = 0
            
            Call Balances.AddBalance(headerRow + 1, "Coinbase", marketCurrency, totalUnits, availableUnits, pendingUnits, accountId)
            Call Dashboard.AddCurrency(marketCurrency, "Coinbase")
                
            marketCurrency = ""
            totalUnits = ""
            availableUnits = ""
            pendingUnits = ""
            publicAddress = ""
            accountId = ""
        Next i
    End If
    
End Sub

Function ParseTrades(sheet As Worksheet, jsonString As String) As Integer

    Application.StatusBar = "Updating Trades - Coinbase"
        
    Dim jsonObject() As String
    Dim resultList() As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim row As Integer
    Dim col As Integer
    Dim found As Boolean
    Dim newTradeCount As Integer
    Dim headerRow As Integer
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
    Dim additionalFees As Double
    
    sheet.Activate
    headerRow = 2
    
    If jsonString = "The service is unavailable." Then
        Exit Function
    End If

    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = UBound(jsonObject) To 0 Step -1
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            obj = Split(resultList(1), """:")
            obj(1) = Replace(obj(1), """", "")
            
            found = False
            For row = headerRow + 1 To sheet.UsedRange.Rows.Count
                If obj(1) = sheet.Cells(row, 1) Then
                    found = True
                    Exit For
                End If
            Next row
            
            If found = False Then
                newTradeCount = newTradeCount + 1
                
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = UCase(Replace(obj(0), """", ""))
                    obj(1) = Replace(obj(1), """", "")
                    If obj(0) = "TRADE_ID" Then
                        id = obj(1)
                    ElseIf obj(0) = "PRODUCT_ID" Then
                        obj = Split(obj(1), "-")
                        marketCurrency = obj(0)
                        baseCurrency = obj(1)
                    ElseIf obj(0) = "TimeStamp" Then
                        openedDate = ISODATE(obj(1))
                    ElseIf obj(0) = "CREATED_AT" Then
                        closedDate = ISODATE(obj(1))
                    ElseIf obj(0) = "SIDE" Then
                        orderType = UCase(obj(1))
                    ElseIf obj(0) = "SIZE" Then
                        units = UCase(obj(1))
                    ElseIf obj(0) = "PRICE" Then
                        rate = UCase(obj(1))
                    ElseIf obj(0) = "FEE" Then
                        commission = UCase(obj(1))
                    ElseIf obj(0) = "USD_VOLUME" Then
                        price = UCase(obj(1))
                    End If
                Next j
                                                
                Call Trades.AddTrade(headerRow + 1, id, "Coinbase", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, CStr(additionalFees))
                Call Portfolio.AddCurrency(marketCurrency, "Coinbase")
                Call Portfolio.AddCurrency(baseCurrency, "Coinbase")
                Call Portfolio.AddMostRecentTrade("Coinbase", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                Call Dashboard.AddCurrency(marketCurrency, "Coinbase")
                Call Dashboard.AddCurrency(baseCurrency, "Coinbase")
                
                id = ""
                baseCurrency = ""
                marketCurrency = ""
                orderType = ""
                units = ""
                rate = ""
                commission = ""
                additionalFees = 0
            End If
        Next i
    End If
    
    ParseTrades = newTradeCount
    
End Function

Sub ParseOrders(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Orders - Coinbase"
           
    Dim jsonObject() As String
    Dim resultList() As String
    Dim str As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim id As String
    Dim exchange As String
    Dim baseCurrency As String
    Dim marketCurrency As String
    Dim orderType As String
    Dim units As String
    Dim limit As String
    Dim openedDate As Date
    
    sheet.Activate
    headerRow = 2
    
    If jsonString = "The service is unavailable." Then
        Exit Sub
    End If
            
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = 0 To UBound(jsonObject)
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
                       
            For j = 0 To UBound(resultList)
                obj = Split(resultList(j), """:")
                obj(0) = UCase(Replace(obj(0), """", ""))
                obj(1) = Replace(obj(1), """", "")
                If obj(0) = "ID" Then
                    id = obj(1)
                ElseIf obj(0) = "PRODUCT_ID" Then
                    obj = Split(obj(1), "-")
                    marketCurrency = obj(0)
                    baseCurrency = obj(1)
                ElseIf obj(0) = "SIDE" Then
                    orderType = UCase(obj(1))
                ElseIf obj(0) = "SIZE" Then
                    units = obj(1)
                ElseIf obj(0) = "PRICE" Then
                    limit = obj(1)
                ElseIf obj(0) = "CREATED_AT" Then
                    openedDate = ISODATE(obj(1))
                End If
            Next j
            
            Call Orders.AddOrder(headerRow + 1, id, "Coinbase", baseCurrency, marketCurrency, orderType, units, limit, CStr(openedDate))
                
            id = ""
            baseCurrency = ""
            marketCurrency = ""
            orderType = ""
            units = ""
            limit = ""
        Next i
    End If
    
End Sub

Sub ParseTransfers(sheet As Worksheet, jsonString As String, coin As String)

    Application.StatusBar = "Updating Transfers - Coinbase"
    
    Dim jsonObject() As String
    Dim resultList() As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim found As Boolean
    Dim units As Double
    Dim fee As Double
    Dim fromDate As Date
    Dim toDate As Date
    Dim transferDate As Date
    Dim txId As String
    Dim fromAcct As String
    Dim toAcct As String
    Dim transactionType As String
    Dim paymentMethodName As String
    Dim subtitle As String
    
    sheet.Activate
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    If jsonString = "The service is unavailable." Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, "data"":")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = UBound(jsonObject) To 0 Step -1
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            jsonObject(i) = Replace(jsonObject(i), "[", "")
            resultList = Split(jsonObject(i), ",")
            
            'Call Helpers.WriteDataToTest(resultList)
            
            obj = Split(resultList(2), """:")
            obj(0) = UCase(Replace(obj(0), """", ""))
            obj(1) = UCase(Replace(obj(1), """", ""))
            
            found = False
            If obj(1) <> "COMPLETED" Then
                found = True
            Else
                obj = Split(resultList(0), """:")
                obj(0) = UCase(Replace(obj(0), """", ""))
                obj(1) = Replace(obj(1), """", "")
                For row = headerRow + 1 To lastRow
                    If obj(1) = sheet.Cells(row, 1) Then
                        found = True
                        Exit For
                    End If
                Next row
            End If
            
            If found = False Then
                obj = Split(resultList(1), """:")
                transactionType = UCase(Replace(obj(UBound(obj)), """", ""))
                
                If transactionType = "SEND" Then
                    If UBound(resultList) = 28 Or UBound(resultList) = 18 Then
                        found = True
                    Else
                        For j = 0 To UBound(resultList)
                            obj = Split(resultList(j), """:")
                            obj(0) = UCase(Replace(obj(0), """", ""))
                            obj(1) = Replace(obj(1), """", "")
                            If obj(0) = "AMOUNT" Then
                                units = Replace(obj(UBound(obj)), """", "")
                                obj = Split(resultList(j + 1), """:")
                                coin = Replace(obj(UBound(obj)), """", "")
                            ElseIf obj(0) = "UPDATED_AT" Then
                                transferDate = ISODATE(obj(1))
                            ElseIf obj(0) = "FEE" Or obj(0) = "TRANSACTION_FEE" Then
                                fee = Replace(obj(UBound(obj)), """", "")
                            ElseIf obj(0) = "SUBTITLE" Then
                                If InStr(obj(1), "From ") > 0 Then
                                    subtitle = Mid(obj(1), 6, Len(obj(1)) - 5)
                                ElseIf InStr(obj(1), "To ") > 0 Then
                                    subtitle = Mid(obj(1), 4, Len(obj(1)) - 3)
                                End If
                            End If
                        Next j
                    End If
                Else
                    For j = 0 To UBound(resultList)
                        obj = Split(resultList(j), """:")
                        obj(0) = UCase(Replace(obj(0), """", ""))
                        obj(1) = Replace(obj(1), """", "")
                        If obj(0) = "TRANSACTION" Then
                            txId = Replace(obj(UBound(obj)), """", "")
                        ElseIf obj(0) = "AMOUNT" Then
                            units = Replace(obj(UBound(obj)), """", "")
                            obj = Split(resultList(j + 1), """:")
                            coin = Replace(obj(UBound(obj)), """", "")
                        ElseIf obj(0) = "UPDATED_AT" Then
                            transferDate = ISODATE(obj(1))
                        ElseIf obj(0) = "FEE" Or obj(0) = "TRANSACTION_FEE" Then
                            fee = Replace(obj(UBound(obj)), """", "")
                        ElseIf obj(0) = "PAYMENT_METHOD_NAME" Then
                            paymentMethodName = Replace(obj(UBound(obj)), """", "")
                        ElseIf obj(0) = "SUBTITLE" Then
                            If InStr(obj(1), "From ") > 0 Then
                                subtitle = Mid(obj(1), 6, Len(obj(1)) - 5)
                            ElseIf InStr(obj(1), "To ") > 0 Then
                                subtitle = Mid(obj(1), 4, Len(obj(1)) - 3)
                            End If
                        End If
                    Next j
                End If
                
                If transactionType = "FIAT_DEPOSIT" Then
                    fromAcct = subtitle
                    toAcct = "Coinbase"
                ElseIf transactionType = "EXCHANGE_DEPOSIT" Then
                    If InStr(obj(1), "From GDAX") > 0 Then
                        fromAcct = ""
                    ElseIf InStr(obj(1), "To GDAX") > 0 Then
                        fromAcct = "Coinbase"
                    End If
                    toAcct = subtitle
                ElseIf transactionType = "EXCHANGE_WITHDRAWAL" Then
                    If InStr(obj(1), "From GDAX") > 0 Then
                        toAcct = "Coinbase"
                    ElseIf InStr(obj(1), "To GDAX") > 0 Then
                        toAcct = ""
                    End If
                    fromAcct = subtitle
                ElseIf transactionType = "FIAT_WITHDRAWAL" Then
                    fromAcct = "Coinbase"
                    toAcct = subtitle
                ElseIf transactionType = "SEND" Then
                    fromAcct = "Coinbase"
                    toAcct = subtitle
                End If
                
                If fromAcct <> "" Then
                    fromDate = transferDate
                End If
                If toAcct <> "" Then
                    toDate = transferDate
                End If
                
                If found = False Then
                    Call Transfers.AddTransfer(fromAcct, toAcct, coin, Abs(units), fee, fromDate, toDate)
                End If
                
                fromAcct = ""
                toAcct = ""
                units = 0
                fee = 0
                fromDate = 0
                toDate = 0
                transferDate = 0
                
            End If
        Next i
    End If
    
End Sub


