Attribute VB_Name = "Trading"
Sub PlaceBuyOrder()
    Set sheet = Sheets("Trading")
    
    Dim exchange As String
    Dim marketCurrency As String
    Dim baseCurrency As String
    Dim quantity As Double
    Dim price As Double
    
    Call DisableApplication
            
    exchange = sheet.Cells(1, 2)
    marketCurrency = UCase(sheet.Cells(2, 2))
    baseCurrency = UCase(sheet.Cells(3, 2))
    quantity = sheet.Cells(15, 2)
    price = sheet.Cells(16, 2)
      
    If exchange = "Bittrex" Then
        Call ApiBittrex.PlaceOrder(baseCurrency & "-" & marketCurrency, quantity, price)
    ElseIf exchange = "Binance" Then
        'Call ApiBinance.PlaceOrder(currencyPair, quantity, price)
    End If
    
    
    Call UpdateOrders
    Call UpdateOrders
    Call EnableApplication
    
End Sub
