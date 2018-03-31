Attribute VB_Name = "Ranges"
Private sheet As Worksheet

Sub UpdateRanges()

    Application.StatusBar = "Updating Ranges"
    
    Set sheet = Sheets("Ranges")
    sheet.Activate
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Cells(headerRow + 1, 1).FormulaR1C1 = "=Trades!RC2"
    sheet.Cells(headerRow + 1, 2).FormulaR1C1 = "=Trades!RC3"
    sheet.Cells(headerRow + 1, 3).FormulaR1C1 = "=Trades!RC4"
    sheet.Cells(headerRow + 1, 4).FormulaR1C1 = "=Trades!RC8*IF(Trades!RC7=""SELL"",-1,1)"
    sheet.Cells(headerRow + 1, 5).FormulaR1C1 = "=Trades!RC9"
    sheet.Columns("A:C").NumberFormat = "General"
    sheet.Range(Cells(headerRow + 1, 1), Cells(headerRow + 1, lastColumn)).Copy
    sheet.Range(Cells(headerRow + 1, 1), Cells(lastRow, lastColumn)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    sheet.Cells(1, 1).Select

    Application.StatusBar = ""
    
End Sub

Sub PlaceOrders()
    Dim PlaceOrder As Boolean
    Dim message As String
    Set sheet = Sheets("Ranges")
    
    Dim exchange As String
    Dim currencyPair As String
    Dim quantity As Double
    Dim price As Double
    Dim orderType As String
    
    Call DisableApplication
    
    Dim r As Range
    For Each r In Selection.Rows
        
        exchange = sheet.Cells(r.row, 1)
        currencyPair = sheet.Cells(r.row, 2) & "-" & sheet.Cells(r.row, 3)
        quantity = sheet.Cells(r.row, 4) * -1
        price = r.Value2
        If quantity < 0 Then
            orderType = "SELL"
        ElseIf quantity > 0 Then
            orderType = "BUY"
        End If
        
        message = "Place " & orderType & " order on " & exchange & " for " & Abs(quantity) & " units of " & currencyPair & " @ " & price & " ?"
        If MsgBox(message, vbYesNo) = vbYes Then
            PlaceOrder = True
        End If
        
        If PlaceOrder = True Then
            If exchange = "Bittrex" Then
                Call ApiBittrex.PlaceOrder(currencyPair, quantity, price)
            ElseIf exchange = "Binance" Then
                Call ApiBinance.PlaceOrder(currencyPair, quantity, price)
            End If
        End If
        
        PlaceOrder = False
    Next r
    
    Call UpdateOrders
    Call EnableApplication
    
End Sub

