Attribute VB_Name = "ApiGDAX"

Function PublicApiGDAX(endpoint As String, Optional MethodOptions As String) As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://api.gdax.com" & endpoint & MethodOptions
    http.Send
    http.WaitForResponse
    PublicApiGDAX = http.ResponseText
    Set http = Nothing

End Function

Function PrivateApiGDAX(Method As String, endpoint As String, Optional options As String) As String
    
    Dim resultList() As String
    Dim url As String
    Dim key As String
    Dim secret As String
    Dim passphrase As String
    Dim timestamp As Long
    Dim requestPath As String
    Dim postdata As String
    Dim signature As String
    
    key = Evaluate(ActiveWorkbook.Names("ApiKeyGDAX").Value)
    secret = Evaluate(ActiveWorkbook.Names("ApiSecretGDAX").Value)
    passphrase = Evaluate(ActiveWorkbook.Names("ApiPassphraseGDAX").Value)
    If Trim(key) = "" Or Trim(secret) = "" Or Trim(passphrase) = "" Then
        Exit Function
    End If
    resultList = Split(PublicApiGDAX("/time", ""), ":")
    timestamp = Replace(resultList(UBound(resultList)), "}", "")
    
    url = "https://api.gdax.com" & endpoint
    postdata = timestamp & UCase(Method) & endpoint & ""
    signature = Base64Encode(ComputeHash_C("SHA256", postdata, Base64Decode(secret), "RAW"))
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open UCase(Method), url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "CB-ACCESS-KEY", key
    http.SetRequestHeader "CB-ACCESS-SIGN", signature
    http.SetRequestHeader "CB-ACCESS-TIMESTAMP", timestamp
    http.SetRequestHeader "CB-ACCESS-PASSPHRASE", passphrase
    http.Send
    http.WaitForResponse
    PrivateApiGDAX = http.ResponseText
    Set http = Nothing

End Function

Sub TestGDAX()
    Dim str As String
    
    'str = PrivateApiGDAX("GET", "/accounts", "")
    
    Call ApiGDAX.ParseBalances(Sheets("Balances"), ApiGDAX.PrivateApiGDAX("GET", "/accounts", ""))
    
    'MsgBox (str)
    
End Sub

Sub ParseBalances(sheet As Worksheet, jsonString As String)

    Application.StatusBar = "Updating Balances - GDAX"
    
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
                    accountId = obj(1)
                ElseIf obj(0) = "CURRENCY" Then
                    marketCurrency = obj(1)
                ElseIf obj(0) = "BALANCE" Then
                    totalUnits = obj(1)
                ElseIf obj(0) = "AVAILABLE" Then
                    availableUnits = obj(1)
                ElseIf obj(0) = "HOLD" Then
                    pendingUnits = obj(1)
                End If
            Next j
            
            Call Balances.AddBalance(headerRow + 1, "GDAX", marketCurrency, totalUnits, availableUnits, pendingUnits, accountId)
            Call Dashboard.AddCurrency(marketCurrency, "GDAX")
                
            marketCurrency = ""
            totalUnits = ""
            availableUnits = ""
            pendingUnits = ""
            accountId = ""
        Next i
    End If
    
End Sub

Function ParseTrades(sheet As Worksheet, jsonString As String) As Integer

    Application.StatusBar = "Updating Trades - GDAX"
        
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
                                
                Call Trades.AddTrade(headerRow + 1, id, "GDAX", baseCurrency, marketCurrency, CStr(openedDate), CStr(closedDate), orderType, units, rate, commission, CStr(additionalFees))
                Call Portfolio.AddCurrency(marketCurrency, "GDAX")
                Call Portfolio.AddCurrency(baseCurrency, "GDAX")
                Call Portfolio.AddMostRecentTrade("GDAX", marketCurrency, closedDate, orderType, CDbl(units), sheet.Cells(headerRow + 1, 15))
                Call Dashboard.AddCurrency(marketCurrency, "GDAX")
                Call Dashboard.AddCurrency(baseCurrency, "GDAX")
                
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

    Application.StatusBar = "Updating Orders - GDAX"
           
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
            
            Call Orders.AddOrder(headerRow + 1, id, "GDAX", baseCurrency, marketCurrency, orderType, units, limit, CStr(openedDate))
                
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

    Application.StatusBar = "Updating Transfers - GDAX"
    
    Dim jsonObject() As String
    Dim resultList() As String
    Dim obj() As String
    Dim i As Integer
    Dim j As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim found As Boolean
    Dim fromAcct As String
    Dim toAcct As String
    Dim units As Double
    Dim fee As Double
    Dim fromDate As Date
    Dim toDate As Date
    Dim transferDate As Date
        
    sheet.Activate
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    If InStr(jsonString, "BadRequest") > 1 Then
        Exit Sub
    End If
    
    jsonObject = Split(jsonString, ",""")
    jsonObject = Split(jsonObject(0), """:")
    If UBound(jsonObject) < 1 Then
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
            
            obj = Split(resultList(4), """:")
            obj(0) = UCase(Replace(obj(0), """", ""))
            obj(1) = UCase(Replace(obj(1), """", ""))
            
            found = False
            If obj(1) <> "TRANSFER" Then
                found = True
            Else
                obj = Split(resultList(5), """:")
                obj(0) = UCase(Replace(obj(0), """", ""))
                obj(1) = UCase(Replace(obj(1), """", ""))
                obj(2) = Replace(obj(2), """", "")
                For row = headerRow + 1 To lastRow
                    If obj(2) = sheet.Cells(row, 1) Then
                        found = True
                        Exit For
                    End If
                Next row
            End If
            
            If found = False Then
                For j = 0 To UBound(resultList)
                    obj = Split(resultList(j), """:")
                    obj(0) = UCase(Replace(obj(0), """", ""))
                    obj(1) = Replace(obj(1), """", "")
                    If obj(0) = "AMOUNT" Then
                        units = obj(1)
                    ElseIf obj(0) = "CREATED_AT" Then
                        transferDate = ISODATE(obj(1))
                    ElseIf obj(0) = "DETAILS" Then
                        txId = Replace(obj(UBound(obj)), """", "")
                    ElseIf obj(0) = "TRANSFER_TYPE" Then
                        If UCase(obj(1)) = "WITHDRAW" Then
                            fromAcct = "GDAX"
                            fromDate = transferDate
                        ElseIf UCase(obj(1)) = "DEPOSIT" Then
                            toAcct = "GDAX"
                            toDate = transferDate
                        End If
                    End If
                Next j
                
                Call Transfers.AddTransfer(fromAcct, toAcct, coin, Abs(units), fee, fromDate, toDate)
                
                toAcct = ""
                fromAcct = ""
                units = 0
                fee = 0
                fromDate = 0
                toDate = 0
                transferDate = 0
                
            End If
        Next i
    End If
    
End Sub

