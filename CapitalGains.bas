Attribute VB_Name = "CapitalGains"
Private sheet As Worksheet

Sub ResetCapitalGains()

    If MsgBox("Are you sure you want to reset capital gains?", vbYesNo) = vbNo Then Exit Sub

    Set sheet = Sheets("Trades")
    sheet.Activate
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 2).End(xlUp).row
    
    sheet.Range(sheet.Cells(headerRow + 1, 16), sheet.Cells(lastRow, 21)).ClearContents
    sheet.Cells(headerRow + 1, 16).FormulaR1C1 = "=IF(RC7=""SELL"","""",RC8)"
    sheet.Cells(headerRow + 1, 17).FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]*RC[-1],"""")"
    sheet.Cells(headerRow + 1, 18).FormulaR1C1 = "=IF(RC7=""BUY"","""",RC13*-1)"
    sheet.Cells(headerRow + 1, 19).FormulaR1C1 = "=IF(RC[-1]<>"""",(RC14/RC13)*RC[-1],"""")"
    sheet.Cells(headerRow + 1, 20).FormulaR1C1 = ""
    sheet.Cells(headerRow + 1, 21).FormulaR1C1 = ""
    
    sheet.Range("P3:U3").Copy
    sheet.Range(sheet.Cells(headerRow + 1, 16), sheet.Cells(lastRow, 21)).Select
    sheet.Paste
    
End Sub

Sub CalculateCapitalGains()

    Application.StatusBar = "Calculating Capital Gains"
    
    Dim curTicker As String
    Dim curExchange As String
    Dim startRow As Integer
    Dim lastRow As Integer
    Dim sellProceeds As Double
    Dim sellDate As Date
    Dim sellUnits As Double
    Dim sellCBPU As Double
    Dim buyUnits As Double
    Dim buyCBPU As Double
    Dim stcg As Double
    Dim ltcg As Double
    Dim startDateHD As Date
    Dim endDateHD As Date
    Dim runCalc As Boolean
    Dim sellRow As Integer
    Dim headerRow As Integer
    
    Set sheet = Sheets("HistoricalQuotes")
    endDateHD = sheet.Cells(sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row, 1)
    startDateHD = sheet.Cells(2, 1)
    If Not IsDate(endDateHD) Or Not IsDate(startDateHD) Then
        Exit Sub
    End If
    
    Set sheet = Sheets("Trades")
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 2).End(xlUp).row
    
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
    
    curExchange = sheet.Cells(lastRow, 2)
    curTicker = sheet.Cells(lastRow, 4)
    startRow = lastRow
    
    For row = lastRow To headerRow + 1 Step -1
        
        If sheet.Cells(row, 3) = "USDT" Or sheet.Cells(row, 4) = "USDT" Then
            row = row
        End If
        
        If sheet.Cells(row, 7) = "BUY" And sheet.Cells(row, 3) = "USD" Then
            'Ignore Buys with a USD base
        ElseIf sheet.Cells(row, 7) = "SELL" And sheet.Cells(row, 20) = "" And sheet.Cells(row, 21) = "" Then
           curExchange = sheet.Cells(row, 2)
           curTicker = sheet.Cells(row, 4)
           sellDate = DateValue(sheet.Cells(row, 6))
           If sellDate >= startDateHD And sellDate <= endDateHD Then
               sellProceeds = (CDbl(sheet.Cells(row, 13).Text) * -1)
               sellUnits = Round(CDbl(sheet.Cells(row, 8).Text), 8)
               sellCBPU = CDbl(sheet.Cells(row, 15).Text) * -1
               stcg = 0
               ltcg = 0
               
               For buyrow = row + 1 To lastRow
                    If curExchange = sheet.Cells(buyrow, 2) And curTicker = sheet.Cells(buyrow, 4) Then
                        If sheet.Cells(buyrow, 7) = "BUY" And sheet.Cells(buyrow, 16) > 0 Then
                            buyUnits = Round(CDbl(sheet.Cells(buyrow, 16).Text), 8)
                            buyCBPU = Round(CDbl(sheet.Cells(buyrow, 15).Text), 8)
                            If sellDate < DateValue(sheet.Cells(buyrow, 6)) + 365 Then
                                stcg = stcg + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            Else
                                ltcg = ltct + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            End If
                            
                            If Round(buyUnits, 8) < Round(sellUnits, 8) Then
                                'get the full cost basis of the buy order and change UQ to 0
                                'continue for loop looking for next buy order
                                sellUnits = Round(sellUnits - buyUnits, 8)
                                sheet.Cells(buyrow, 16) = 0
                            Else
                                'calculate and populate CG field
                                'update UQ field to subtract off SELL Quantity
                                sheet.Cells(row, 20) = Round(stcg, 2)
                                sheet.Cells(row, 21) = Round(ltcg, 2)
                                sheet.Cells(buyrow, 16) = Round(buyUnits - sellUnits, 8)
                                Exit For
                            End If
                        End If
                    ElseIf curExchange = sheet.Cells(buyrow, 2) And curTicker = sheet.Cells(buyrow, 3) Then
                        If sheet.Cells(buyrow, 7) = "SELL" And sheet.Cells(buyrow, 18) > 0 Then
                            buyUnits = Round(CDbl(sheet.Cells(buyrow, 18).Text), 8)
                            buyCBPU = Round(CDbl(sheet.Cells(buyrow, 14).Text) / CDbl(sheet.Cells(buyrow, 13).Text), 8)
                            If sellDate < DateValue(sheet.Cells(buyrow, 6)) + 365 Then
                                stcg = stcg + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            Else
                                ltcg = ltct + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            End If
                            
                            If buyUnits < sellUnits Then
                                'get the full cost basis of the buy order and change UQ to 0
                                'continue for loop looking for next buy order
                                sellUnits = Round(sellUnits - buyUnits, 8)
                                sheet.Cells(buyrow, 18) = 0
                            Else
                                'calculate and populate CG field
                                'update UQ field to subtract off SELL Quantity
                                sheet.Cells(row, 20) = Round(stcg, 2)
                                sheet.Cells(row, 21) = Round(ltcg, 2)
                                sheet.Cells(buyrow, 18) = Round(buyUnits - sellUnits, 8)
                                Exit For
                            End If
                        End If
                    End If
               Next buyrow
           End If
        ElseIf sheet.Cells(row, 7) = "BUY" And sheet.Cells(row, 20) = "" And sheet.Cells(row, 21) = "" Then
            curExchange = sheet.Cells(row, 2)
            curTicker = sheet.Cells(row, 3)
            sellDate = DateValue(sheet.Cells(row, 6))
            If sellDate >= startDateHD And sellDate <= endDateHD Then
                sellProceeds = Round(CDbl(sheet.Cells(row, 14).Text), 8)
                sellUnits = Round(CDbl(sheet.Cells(row, 13).Text), 8)
                sellCBPU = Round(sellProceeds / sellUnits, 8)
                stcg = 0
                ltcg = 0
               
               For buyrow = row + 1 To lastRow
                    If curExchange = sheet.Cells(buyrow, 2) And curTicker = sheet.Cells(buyrow, 4) Then
                        If sheet.Cells(buyrow, 7) = "BUY" And sheet.Cells(buyrow, 16) > 0 Then
                            buyUnits = Round(CDbl(sheet.Cells(buyrow, 16).Text), 8)
                            buyCBPU = Round(CDbl(sheet.Cells(buyrow, 15).Text), 8)
                            If sellDate < DateValue(sheet.Cells(buyrow, 6)) + 365 Then
                                stcg = stcg + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            Else
                                ltcg = ltct + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            End If
                            
                            If buyUnits < sellUnits Then
                                'get the full cost basis of the buy order and change UQ to 0
                                'continue for loop looking for next buy order
                                sellUnits = Round(sellUnits - buyUnits, 8)
                                sheet.Cells(buyrow, 16) = 0
                            Else
                                'calculate and populate CG field
                                'update UQ field to subtract off SELL Quantity
                                sheet.Cells(row, 20) = Round(stcg, 2)
                                sheet.Cells(row, 21) = Round(ltcg, 2)
                                sheet.Cells(buyrow, 16) = Round(buyUnits - sellUnits, 8)
                                Exit For
                            End If
                        End If
                    ElseIf curExchange = sheet.Cells(buyrow, 2) And curTicker = sheet.Cells(buyrow, 3) Then
                        If sheet.Cells(buyrow, 7) = "SELL" And sheet.Cells(buyrow, 18) > 0 Then
                            buyUnits = Round(CDbl(sheet.Cells(buyrow, 18).Text), 8)
                            buyCBPU = Round(CDbl(sheet.Cells(buyrow, 14).Text) / CDbl(sheet.Cells(buyrow, 13).Text), 8)
                            If sellDate < DateValue(sheet.Cells(buyrow, 6)) + 365 Then
                                stcg = stcg + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            Else
                                ltcg = ltct + (WorksheetFunction.Min(sellUnits, buyUnits) * sellCBPU) - (WorksheetFunction.Min(sellUnits, buyUnits) * buyCBPU)
                            End If
                            
                            If buyUnits < sellUnits Then
                                'get the full cost basis of the buy order and change UQ to 0
                                'continue for loop looking for next buy order
                                sellUnits = Round(sellUnits - buyUnits, 8)
                                sheet.Cells(buyrow, 18) = 0
                            Else
                                'calculate and populate CG field
                                'update UQ field to subtract off SELL Quantity
                                sheet.Cells(row, 20) = Round(stcg, 2)
                                sheet.Cells(row, 21) = Round(ltcg, 2)
                                sheet.Cells(buyrow, 18) = Round(buyUnits - sellUnits, 8)
                                Exit For
                            End If
                        End If
                    End If
               Next buyrow
           End If
        End If
    Next row
    
    Application.StatusBar = ""
    
End Sub

