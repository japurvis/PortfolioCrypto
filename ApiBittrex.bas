Attribute VB_Name = "ApiBittrex"

Function PublicApiBittrex(Method As String, Optional MethodOptions As String) As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://bittrex.com/api/v1.1/public/" & Method & MethodOptions
    http.Send
    http.WaitForResponse
    PublicApi = http.ResponseText
    Set http = Nothing

End Function

Function PrivateApiBittrex(Method As String, Optional MethodOptions As String) As String
    
    Dim apikey As String
    Dim apisecret As String
    Dim nonce As String
    
    apikey = Evaluate(ActiveWorkbook.Names("ApiKeyBittrex").Value)
    apisecret = Evaluate(ActiveWorkbook.Names("ApiSecretBittrex").Value)
    
    If Trim(apikey) = "" Or Trim(apisecret) = "" Then
        Exit Function
    End If
    
    nonce = DateDiff("s", "1/1/1970", now)
    
    apiUrl = "https://bittrex.com/api/v1.1/"
    postdata = Method & "?apikey=" & apikey & MethodOptions & "&nonce=" & nonce
    APIsign = ComputeHash_C("SHA512", apiUrl & postdata, apisecret, "STRHEX")
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", apiUrl & postdata, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.SetRequestHeader "apisign", APIsign
    http.Send (postdata)
    http.WaitForResponse
    PrivateApiBittrex = http.ResponseText
    Set http = Nothing

End Function

Sub CancelOrder(orderId As String)
    
    Call PrivateApiBittrex("market/cancel", "&uuid=" & orderId)
    
End Sub

Function PlaceOrder(baseCurrency As String, marketCurrency As String, quantity As Double, rate As Double) As String
    
    If quantity > 0 Then
        PlaceOrder = PlaceBuyOrder(baseCurency, marketCurrency, quantity, rate)
    ElseIf quantity < 0 Then
        PlaceOrder = PlaceSellOrder(baseCurency, marketCurrency, Abs(quantity), rate)
    End If
    
End Function

Function PlaceBuyOrder(baseCurrency As String, marketCurrency As String, quantity As Double, rate As Double) As String
    If marketCurrency = "BCH" Then
        marketCurrency = "BCC"
    End If
    
    PlaceBuyOrder = PrivateApiBittrex("market/buylimit", "&market=" & baseCurrency & "-" & marketCurrency & "&quantity=" & quantity & "&rate=" & rate)
    
End Function

Function PlaceSellOrder(baseCurrency As String, marketCurrency As String, quantity As Double, rate As Double) As String
    If marketCurrency = "BCH" Then
        marketCurrency = "BCC"
    End If
    
    PlaceSellOrder = PrivateApiBittrex("market/selllimit", "&market=" & baseCurrency & "-" & marketCurrency & "&quantity=" & quantity & "&rate=" & rate)
    
End Function

Sub ParseBalances(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Balances - Bittrex"
    
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
    
    If jsonString = "The service is unavailable." Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
        Exit Sub
    End If
    If jsonObject(1) <> "true" Then
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
            
            obj = Split(resultList(0), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
            
            For j = 0 To UBound(resultList)
                obj = Split(resultList(j), """:")
                obj(0) = Replace(obj(0), """", "")
                obj(1) = Replace(obj(1), """", "")
                If (obj(1) <> "null" And obj(1) <> "NONE") Then
                    If (obj(1) <> "null") Then
                        If UCase(obj(0)) = "CURRENCY" Then
                            If UCase(obj(1)) = "BCC" Then
                                marketCurrency = "BCH"
                            Else
                                marketCurrency = obj(1)
                            End If
                        ElseIf UCase(obj(0)) = "BALANCE" Then
                            totalUnits = obj(1)
                        ElseIf UCase(obj(0)) = "AVAILABLE" Then
                            availableUnits = obj(1)
                        ElseIf UCase(obj(0)) = "PENDING" Then
                            pendingUnits = obj(1)
                        End If
                    End If
                End If
            Next j
            
            Call Balances.AddBalance(headerRow + 1, "Bittrex", marketCurrency, totalUnits, availableUnits, pendingUnits, "")
            Call Dashboard.AddCurrency(marketCurrency, "Bittrex")
                
            marketCurrency = ""
            totalUnits = ""
            availableUnits = ""
            pendingUnits = ""
        Next i
    End If
    
End Sub

Sub ParseOrders(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Orders - Bittrex"
           
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
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
        Exit Sub
    End If
    If jsonObject(1) <> "true" Then
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
            
            obj = Split(resultList(0), """:")
            obj(0) = Replace(obj(0), """", "")
            obj(1) = Replace(obj(1), """", "")
            
            For j = 0 To UBound(resultList)
                obj = Split(resultList(j), """:")
                obj(0) = Replace(obj(0), """", "")
                obj(1) = Replace(obj(1), """", "")
                If (obj(1) <> "null" And obj(1) <> "NONE") Then
                    If (obj(1) <> "null") Then
                        If obj(0) = "OrderUuid" Then
                            id = obj(1)
                        ElseIf obj(0) = "Exchange" Then
                            obj = Split(obj(1), "-")
                            baseCurrency = obj(0)
                            If UCase(obj(1)) = "BCC" Then
                                marketCurrency = "BCH"
                            Else
                                marketCurrency = obj(1)
                            End If
                        ElseIf obj(0) = "OrderType" Then
                            obj = Split(obj(1), "_")
                            orderType = obj(UBound(obj))
                        ElseIf obj(0) = "QuantityRemaining" Then
                            units = obj(1)
                        ElseIf obj(0) = "Limit" Then
                            limit = obj(1)
                        ElseIf obj(0) = "Opened" Then
                            openedDate = ISODATE(obj(1))
                        ElseIf obj(0) = "Condition" Then
                        ElseIf obj(0) = "ConditionTarget" Then
                        End If
                    End If
                End If
            Next j
                        
            Call Orders.AddOrder(headerRow + 1, id, "Bittrex", baseCurrency, marketCurrency, orderType, units, limit, CStr(openedDate))
                
            id = ""
            baseCurrency = ""
            marketCurrency = ""
            orderType = ""
            units = ""
            limit = ""
        Next i
    End If
    
End Sub

Function ParseTrades(sheet As Worksheet, jsonString As String) As Integer

    Application.StatusBar = "Updating Trades - Bittrex"
        
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

    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
        Exit Function
    End If
    If jsonObject(1) <> "true" Then
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
            
            obj = Split(resultList(0), """:")
            obj(0) = Replace(obj(0), """", "")
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
                    obj(0) = Replace(obj(0), """", "")
                    obj(1) = Replace(obj(1), """", "")
                    If (obj(1) <> "null" And obj(1) <> "NONE") Then
                        If obj(0) = "OrderUuid" Then
                            id = obj(1)
                        ElseIf obj(0) = "Exchange" Then
                            obj = Split(obj(1), "-")
                            baseCurrency = obj(0)
                            If UCase(obj(1)) = "BCC" Then
                                marketCurrency = "BCH"
                            Else
                                marketCurrency = obj(1)
                            End If
                        ElseIf obj(0) = "TimeStamp" Then
                            openedDate = ISODATE(obj(1))
                        ElseIf obj(0) = "Closed" Then
                            closedDate = ISODATE(obj(1))
                        ElseIf obj(0) = "OrderType" Then
                            obj = Split(obj(1), "_")
                            orderType = obj(UBound(obj))
                        ElseIf obj(0) = "Quantity" Then
                            units = UCase(obj(1))
                        ElseIf obj(0) = "PricePerUnit" Then
                            rate = UCase(obj(1))
                        ElseIf obj(0) = "Commission" Then
                            commission = UCase(obj(1))
                        ElseIf obj(0) = "Price" Then
                            price = UCase(obj(1))
                        End If
                    End If
                Next j
                
                additionalFees = price - (units * rate)
                
                Call Trades.AddTrade(headerRow + 1, id, "Bittrex", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, CStr(additionalFees))
                Call Portfolio.AddCurrency(marketCurrency, "Bittrex")
                Call Portfolio.AddCurrency(baseCurrency, "Bittrex")
                Call Portfolio.AddMostRecentTrade("Bittrex", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                Call Dashboard.AddCurrency(marketCurrency, "Bittrex")
                Call Dashboard.AddCurrency(baseCurrency, "Bittrex")
                
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

Sub ParseTransfers(sheet As Worksheet, jsonString As String, fromAcct As String, toAcct As String)

    Application.StatusBar = "Updating Transfers - Bittrex"
    
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
    
    If jsonString = "The service is unavailable." Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
        Exit Sub
    End If
    If jsonObject(1) <> "true" Then
        Exit Sub
    End If
          
    jsonObject = Split(jsonString, "[")
    jsonObject = Split(jsonObject(1), "},{")
    
    If InStr(jsonObject(0), ":") > 0 Then
        For i = UBound(jsonObject) To 0 Step -1
            jsonObject(i) = Replace(jsonObject(i), "}", "")
            jsonObject(i) = Replace(jsonObject(i), "{", "")
            jsonObject(i) = Replace(jsonObject(i), "]", "")
            resultList = Split(jsonObject(i), ",")
            
            'Call Helpers.WriteDataToTest(resultList)
            
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
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = UCase(Replace(obj(0), """", ""))
                    obj(1) = Replace(obj(1), """", "")
                    If (obj(1) <> "null") Then
                        If obj(0) = "CURRENCY" Then
                            If obj(1) = "BCC" Then
                                coin = "BCH"
                            Else
                                coin = obj(1)
                            End If
                        ElseIf obj(0) = "AMOUNT" Then
                            units = obj(1)
                        ElseIf obj(0) = "TXCOST" Then
                            fee = obj(1)
                        ElseIf obj(0) = "OPENED" Or obj(0) = "LASTUPDATED" Then
                            transferDate = ISODATE(obj(1))
                        ElseIf obj(0) = "CANCELED" Then
                            canceled = UCase(obj(1))
                        End If
                    End If
                Next j
                
                If canceled <> "TRUE" Then
                    If UCase(fromAcct) = "BITTREX" Then
                        fromDate = transferDate
                    ElseIf UCase(toAcct) = "BITTREX" Then
                        toDate = transferDate
                    End If
                        
                    Call Transfers.AddTransfer(fromAcct, toAcct, coin, Abs(units), fee, fromDate, toDate)
                End If
                
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
