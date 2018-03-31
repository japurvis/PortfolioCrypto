Attribute VB_Name = "Balances"
Private sheet As Worksheet

Sub UpdateBalancesSheet()
    
    Call DisableApplication
    Call UpdateBalances
    Call EnableApplication
    
End Sub

Sub UpdateBalances()
    
    Application.StatusBar = "Updating Balances"
    
    Call PreFormat
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBittrex").Value) = 1 Then
        Call ApiBittrex.ParseBalances(Sheets("Balances"), PrivateApiBittrex("account/getbalances"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataBinance").Value) = 1 Then
        Call ApiBinance.ParseBalances(Sheets("Balances"), ApiBinance.PrivateApiBinance("GET", "account"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataGDAX").Value) = 1 Then
        Call ApiGDAX.ParseBalances(Sheets("Balances"), ApiGDAX.PrivateApiGDAX("GET", "/accounts"))
    End If
    
    If Evaluate(ActiveWorkbook.Names("ApiLoadDataCoinbase").Value) = 1 Then
        Call ApiCoinbase.ParseBalances(Sheets("Balances"), ApiCoinbase.PrivateApiCoinbase("GET", "/accounts", "?&limit=100"))
    End If
    
    Call PostFormat
    
    Application.StatusBar = ""
    
End Sub

Sub AddBalance(row As Integer, exchange As String, marketCurrency As String, totalUnits As String, availableUnits As String, pendingUnits As String, accountId As String)

    Set sheet = Sheets("Balances")
    Dim col As Integer
    col = 1
    
    sheet.Rows(row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        
    sheet.Cells(row, col) = exchange & "-" & marketCurrency
    col = col + 1
    sheet.Cells(row, col) = exchange
    col = col + 1
    sheet.Cells(row, col) = marketCurrency
    col = col + 1
    sheet.Cells(row, col) = totalUnits
    col = col + 1
    sheet.Cells(row, col) = availableUnits
    col = col + 1
    sheet.Cells(row, col) = pendingUnits
    col = col + 1
    sheet.Cells(row, col) = accountId
            
End Sub

Private Sub PreFormat()
    
    Set sheet = Sheets("Balances")
    sheet.Activate
        
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.UsedRange.Rows.Count
    lastColumn = sheet.UsedRange.Columns.Count
    
    If lastRow > headerRow Then
        Rows("" & headerRow + 1 & ":" & lastRow & "").Delete Shift:=xlUp
    End If
    
End Sub

Private Sub PostFormat()
    
    Set sheet = Sheets("Balances")
    sheet.Activate
    
    Dim headerRow As Integer
    Dim lastRow As Integer
    Dim lastColumn As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
    lastColumn = sheet.Cells(headerRow, sheet.UsedRange.Columns.Count + 1).End(xlToLeft).Column
    
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add key:=Range(sheet.Cells(headerRow + 1, 1), sheet.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
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

Function GetBalanceCollection() As Collection

    Dim curSheet As Worksheet
    Set curSheet = ActiveSheet
    
    Set sheet = Sheets("Balances")
    sheet.Activate
    
    Dim coinList As Collection
    Set coinList = New Collection
      
    Dim headerRow As Integer
    Dim lastRow As Integer
    
    headerRow = 2
    lastRow = sheet.Cells(sheet.UsedRange.Rows.Count + 1, 1).End(xlUp).row
        
    For row = headerRow + 1 To lastRow
        coinList.Add (sheet.Cells(row, 2) & "|" & sheet.Cells(row, 3) & "|" & sheet.Cells(row, 4))
    Next row
    
    Set GetBalanceCollection = coinList
    
    curSheet.Activate
End Function

