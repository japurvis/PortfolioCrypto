Attribute VB_Name = "Portfolio"
Private sheet As Worksheet

Sub UpdatePortfolioSheet()
    
    Call DisableApplication
    Call UpdatePortfolio
    Call EnableApplication
    
End Sub

Sub UpdatePortfolio()

    Application.StatusBar = "Updating Portfolio"
    
    Call PreFormat
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim coinList As Collection
    Set coinList = GetBalanceCollection
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 11).End(xlUp).row
    
    Dim coinFound As Boolean
    Dim exchange As String
    Dim coin As String
    Dim units As String
    If (coinList.Count > 0) Then
        
        For i = 1 To coinList.Count
            coinFound = False
            exchange = Split(coinList(i), "|")(0)
            coin = Split(coinList(i), "|")(1)
            units = Split(coinList(i), "|")(2)
        
            For row = headerRow + 1 To lastRow - 1
                If sheet.Cells(row, 1) = exchange And sheet.Cells(row, 2) = coin Then
                    coinFound = True
                    If units = 0 Then
                        sheet.Rows(row).Delete Shift:=xlUp
                        lastRow = lastRow - 1
                    End If
                    Exit For
                End If
            Next row
            
            If coinFound = False And units > 0 Then
                Call AddCurrency(coin, exchange)
                lastRow = lastRow + 1
            End If
        Next i
    End If
    
    Call PostFormat
    
    Application.StatusBar = ""
    
End Sub

Sub AddCurrency(coin As String, exchange As String)

    If coin = "" Or exchange = "" Then
        Exit Sub
    End If
    
    Set sheet = Sheets("Portfolio")
    sheet.Activate
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    For row = headerRow + 1 To lastRow - 1
        If sheet.Cells(row, 1) = exchange And sheet.Cells(row, 2) = coin Then
            coinFound = True
            Exit For
        End If
    Next row
    
    If coinFound = True Then
        Exit Sub
    End If
    
    sheet.Rows(headerRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    sheet.Cells(headerRow + 1, 1).FormulaR1C1 = exchange
    sheet.Cells(headerRow + 1, 2).FormulaR1C1 = UCase(coin)
    If coin = "USD" Then
        sheet.Cells(headerRow + 1, 3) = "United States Dollar"
        sheet.Cells(headerRow + 1, 4).FormulaR1C1 = "=VLOOKUP(RC1&""-""&RC2,Balances,4,FALSE)"
        sheet.Cells(headerRow + 1, 5) = 0
        sheet.Cells(headerRow + 1, 7) = 1
    Else
        sheet.Cells(headerRow + 1, 3).FormulaR1C1 = "=IFERROR(VLOOKUP(RC1&""-BTC-""&RC2,Quotes,6,FALSE),IFERROR(VLOOKUP(RC1&""-USDT-""&RC2,Quotes,6,FALSE),IFERROR(VLOOKUP(RC1&""-USD-""&RC2,Quotes,6,FALSE),IF(RC2=""USDT"",VLOOKUP(""KRAKEN-USD-USDT"",Quotes,6,FALSE),""""))))"
        sheet.Cells(headerRow + 1, 4).FormulaR1C1 = "=SUMIFS(Trades!C8,Trades!C4,RC2,Trades!C7,""BUY"",Trades!C2,RC1)-SUMIFS(Trades!C8,Trades!C4,RC2,Trades!C7,""SELL"",Trades!C2,RC1)-SUMIFS(Trades!C13,Trades!C3,RC2,Trades!C7,""BUY"",Trades!C2,RC1)-SUMIFS(Trades!C13,Trades!C3,RC2,Trades!C7,""SELL"",Trades!C2,RC1)"
        sheet.Cells(headerRow + 1, 5).FormulaR1C1 = "=IFERROR(SUMIFS(Trades!C17,Trades!C4,RC2,Trades!C2,RC1)+SUMIFS(Trades!C19,Trades!C3,RC2,Trades!C2,RC1),"""")"
        sheet.Cells(headerRow + 1, 7).FormulaR1C1 = "=IF(RC2=""USDT"",VLOOKUP(""Kraken-USD-USDT"",Quotes,7,FALSE),IFERROR(VLOOKUP(RC1&""-USD-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC1&""-USDT-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC1&""-BTC-""&RC2,Quotes,7,FALSE),0)*MAX(IFERROR(VLOOKUP(RC1&""-USDT-BTC"",Quotes,7,FALSE),0),IFERROR(VLOOKUP(RC1&""-USD-BTC"",Quotes,7,FALSE),0)))))"
    End If
    
    sheet.Cells(headerRow + 1, 6).FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],"""")"
    sheet.Cells(headerRow + 1, 8).FormulaR1C1 = "=IFERROR(RC4*RC7,"""")"
    sheet.Cells(headerRow + 1, 9).FormulaR1C1 = "=IFERROR(RC8-RC5,"""")"
    sheet.Cells(headerRow + 1, 10).FormulaR1C1 = "=IFERROR((RC7-RC6)/RC6,0)"
    sheet.Cells(headerRow + 1, 11).FormulaR1C1 = "=IFERROR(RC8/R" & lastRow + 1 & "C8,"""")"
       
    sheet.Cells(headerRow + 1, 13).ClearContents
    sheet.Cells(headerRow + 1, 14).FormulaR1C1 = "=IFERROR(IF(ABS((RC11-RC[-1])/RC[-1])>TargetThreshold,(PortfolioMarketValue*RC[-1]/RC7)-RC4,""""),"""")"
    sheet.Cells(headerRow + 1, 19).FormulaR1C1 = "=IFERROR((RC7-RC[-1])*RC[-2],"""")"
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 2), sheet.Cells(lastRow, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sheet.Sort
        .SetRange Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Private Sub PreFormat()

    Set sheet = Sheets("Portfolio")
    sheet.Activate
        
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 11).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    If lastRow > headerRow + 1 Then
        sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow - 1, lastColumn)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    End If
    
End Sub

Private Sub PostFormat()
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Cells(lastRow, 5).Formula = "=SUM($E" & headerRow + 1 & ":$E" & lastRow - 1 & ")"
    sheet.Cells(lastRow, 8).Formula = "=SUM($H" & headerRow + 1 & ":$H" & lastRow - 1 & ")"
    sheet.Cells(lastRow, 9).Formula = "=SUM($I" & headerRow + 1 & ":$I" & lastRow - 1 & ")"
    sheet.Cells(lastRow, 10).FormulaR1C1 = "=IFERROR(RC9/RC5,0)"
    sheet.Cells(lastRow, 11).Formula = "=SUM($K" & headerRow + 1 & ":$K" & lastRow - 1 & ")"
    sheet.Cells(lastRow, 13).Formula = "=SUM($M" & headerRow + 1 & ":$M" & lastRow - 1 & ")"
    sheet.Cells(lastRow, 19).Formula = "=SUM($S" & headerRow + 1 & ":$S" & lastRow - 1 & ")"
    
    sheet.Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow - 1, 3)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    sheet.Range(sheet.Cells(headerRow + 1, 4), sheet.Cells(lastRow - 1, lastColumn)).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    sheet.Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow - 1, 1)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=API!$A:$A"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow - 1, 11)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 13), sheet.Cells(lastRow - 1, 14)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 16), sheet.Cells(lastRow - 1, 19)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow - 1, lastColumn)).Font.Bold = True
    sheet.Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow, lastColumn)).EntireRow.AutoFit
    sheet.Cells(1, 1).Select
    
    Dim coinList As Collection
    Set coinList = GetBalanceCollection
    
    If (coinList.Count > 0) Then
        Dim exchange As String
        Dim coin As String
        Dim units As Double
        Dim i As Integer
        Dim row As Integer
        
        For i = 1 To coinList.Count
            coinFound = False
            exchange = Split(coinList(i), "|")(0)
            coin = Split(coinList(i), "|")(1)
            units = Round(Split(coinList(i), "|")(2), 8)
        
            For row = headerRow + 1 To lastRow - 1
                If sheet.Cells(row, 1) = exchange And sheet.Cells(row, 2) = coin Then
                    sheet.Cells(row, 4).Select
                    If units <> Round(sheet.Cells(row, 4), 8) Then
                        With Selection.Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    Else
                        With Selection.Font
                            .ColorIndex = xlAutomatic
                            .TintAndShade = 0
                        End With
                    End If
                    Exit For
                End If
            Next row
        Next i
    End If
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 2), sheet.Cells(lastRow, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sheet.Sort
        .SetRange Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastColumn))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Sub AddMostRecentTrade(exchange As String, coin As String, tradeDate As Date, orderType As String, units As Double, price As Currency)
    Set sheet = Sheets("Portfolio")
    sheet.Activate
    
    Dim row As Integer
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 11).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    If orderType = "SELL" And units > 0 Then
        units = units * -1
    End If
    
    If price < 0 Then
        price = price * -1
    End If
    
    For row = headerRow + 1 To lastRow
        If sheet.Cells(row, 1) = exchange And sheet.Cells(row, 2) = coin Then
            If sheet.Cells(row, 16) = "" Or sheet.Cells(row, 16) < tradeDate Then
                sheet.Cells(row, 16) = tradeDate
                sheet.Cells(row, 17) = units
                sheet.Cells(row, 18) = price
                sheet.Cells(row, 19).FormulaR1C1 = "=IFERROR((RC7-RC[-1])*RC[-2],"""")"
            End If
            Exit For
        End If
    Next row
End Sub


