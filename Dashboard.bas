Attribute VB_Name = "Dashboard"
Private sheet As Worksheet

Sub Update()

    Application.StatusBar = "Updating"
    
    Call DisableApplication
        
    Dim newTradeCount As Integer
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    
    Call UpdateQuotes
    Call UpdateHistoricalQuotes
    Call UpdateBalances
    Call UpdateOrders
    newTradeCount = UpdateTrades
    Call UpdateTransfers
    Call UpdateRanges
    Call UpdatePortfolio
    
    Sheets("Dashboard").Cells(1, 2) = now
    Sheets("Dashboard").Cells(2, 2) = newTradeCount
    Call PostFormat
    Application.StatusBar = ""
    
    currentSheet.Activate
    Call EnableApplication
    
End Sub

Private Sub PostFormat()
    
    Set sheet = Sheets("Dashboard")
    sheet.Activate
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 6
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    
    sheet.Range("A" & headerRow + 2 & ":M" & lastRow).Select
    ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Add key:=sheet.Range("A" & headerRow + 2 & ":A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Add key:=sheet.Range("B" & headerRow + 2 & ":B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Dashboard").Sort
        .SetRange sheet.Range("A" & headerRow + 2 & ":M" & lastRow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, 9)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 11), sheet.Cells(lastRow, 13)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
    sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, 13)).EntireColumn.AutoFit
    sheet.Cells(1, 1).Select
    
End Sub


Sub AddCurrency(coin As String, exchange As String)
    
    If coin = "" Or coin = "USD" Then
        Exit Sub
    End If

    Set sheet = Sheets("Dashboard")
    sheet.Activate
    
    Dim bool As Boolean
    Dim headerRow As Integer
    Dim totalRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 6
    totalRow = headerRow + 1
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    found = False
    
    For row = totalRow + 1 To lastRow
        If sheet.Cells(row, 1) = exchange And sheet.Cells(row, 2) = coin Then
            found = True
            Exit For
        End If
    Next row
        
    If found = False Then
        sheet.Rows(totalRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        
        sheet.Cells(totalRow + 1, 1) = exchange
        sheet.Cells(totalRow + 1, 2) = UCase(coin)
        
        sheet.Cells(totalRow + 1, 3).FormulaR1C1 = "=SUMIFS(Trades!C20,Trades!C2,RC1,Trades!C4,RC2,Trades!C[4],""SELL"")+SUMIFS(Trades!C20,Trades!C2,RC1,Trades!C3,RC2,Trades!C[4],""BUY"")"
        sheet.Cells(totalRow + 1, 4).FormulaR1C1 = "=SUMIFS(Trades!C21,Trades!C2,RC1,Trades!C4,RC2,Trades!C[3],""SELL"")+SUMIFS(Trades!C21,Trades!C2,RC1,Trades!C3,RC2,Trades!C[3],""BUY"")"
        sheet.Cells(totalRow + 1, 5).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"

        sheet.Cells(totalRow + 1, 6).FormulaR1C1 = _
            "=((SUMIFS(Trades!C[2],Trades!C2,RC1,Trades!C4,RC2,Trades!C,"">""&OneYearAgo,Trades!C[1],""BUY"")-SUMIFS(Trades!C[2],Trades!C2,RC1,Trades!C4,RC2,Trades!C,"">""&OneYearAgo,Trades!C[1],""SELL"")-SUMIFS(Trades!C[7],Trades!C2,RC1,Trades!C3,RC2,Trades!C,"">""&OneYearAgo,Trades!C[1],""BUY"")-SUMIFS(Trades!C[7],Trades!C2,RC1,Trades!C3,RC2,Trades!C,"">""&OneYearAgo,Trades!C[" & _
            "1],""SELL""))*(IF(RC2=""USDT"",VLOOKUP(""Kraken-USD-USDT"",Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-5]&""-USD-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-5]&""-USDT-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-5]&""-BTC-""&RC2,Quotes,7,FALSE),0)*MAX(IFERROR(VLOOKUP(RC[-5]&""-USDT-BTC"",Quotes,7,FALSE),0),IFERROR(VLOOKUP(RC[-5]&""-USD-BTC"",Quotes,7,FALSE),0)))))))-(IFERROR(" & _
            "SUMIFS(Trades!C[11],Trades!C2,RC1,Trades!C4,RC2,Trades!C,"">""&OneYearAgo)+SUMIFS(Trades!C[13],Trades!C2,RC1,Trades!C3,RC2,Trades!C,"">""&OneYearAgo),""""))" & _
            ""
        sheet.Cells(totalRow + 1, 7).FormulaR1C1 = _
            "=((SUMIFS(Trades!C[1],Trades!C2,RC1,Trades!C4,RC2,Trades!C[-1],""<=""&OneYearAgo,Trades!C,""BUY"")-SUMIFS(Trades!C[1],Trades!C2,RC1,Trades!C4,RC2,Trades!C[-1],""<=""&OneYearAgo,Trades!C,""SELL"")-SUMIFS(Trades!C[6],Trades!C2,RC1,Trades!C3,RC2,Trades!C[-1],""<=""&OneYearAgo,Trades!C,""BUY"")-SUMIFS(Trades!C[6],Trades!C2,RC1,Trades!C3,RC2,Trades!C[-1],""<=""&OneYearAg" & _
            "o,Trades!C,""SELL""))*(IF(RC2=""USDT"",VLOOKUP(""Kraken-USD-USDT"",Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-6]&""-USD-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-6]&""-USDT-""&RC2,Quotes,7,FALSE),IFERROR(VLOOKUP(RC[-6]&""-BTC-""&RC2,Quotes,7,FALSE),0)*MAX(IFERROR(VLOOKUP(RC[-6]&""-USDT-BTC"",Quotes,7,FALSE),0),IFERROR(VLOOKUP(RC[-6]&""-USD-BTC"",Quotes,7,FALSE),0)))))))-(" & _
            "IFERROR(SUMIFS(Trades!C[10],Trades!C2,RC1,Trades!C4,RC2,Trades!C[-1],""<=""&OneYearAgo)+SUMIFS(Trades!C[12],Trades!C2,RC1,Trades!C3,RC2,Trades!C[-1],""<=""&OneYearAgo),""""))" & _
            ""
        sheet.Cells(totalRow + 1, 8).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
        sheet.Cells(totalRow + 1, 9).FormulaR1C1 = "=RC[-4]+RC[-1]"
        sheet.Cells(totalRow + 1, 11).FormulaR1C1 = "=COUNTIFS(Orders!C2,RC1,Orders!C4,RC2,Orders!C5,R" & headerRow & "C)"
        sheet.Cells(totalRow + 1, 12).FormulaR1C1 = "=COUNTIFS(Orders!C2,RC1,Orders!C4,RC2,Orders!C5,R" & headerRow & "C)"
        sheet.Cells(totalRow + 1, 13).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
        
        sheet.Cells(totalRow, 3).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 4).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 5).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 6).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 7).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 8).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 9).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 11).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 12).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        sheet.Cells(totalRow, 13).FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow + 1 - totalRow & "]C)"
        
        sheet.Range("A" & totalRow + 1 & ":M" & lastRow + 1).Select
        ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Add key:=sheet.Range("A" & totalRow + 1 & ":A" & lastRow + 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Dashboard").Sort.SortFields.Add key:=sheet.Range("B" & totalRow + 1 & ":B" & lastRow + 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Dashboard").Sort
            .SetRange sheet.Range("A" & totalRow + 1 & ":M" & lastRow + 1)
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
End Sub


