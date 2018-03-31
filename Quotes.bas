Attribute VB_Name = "Quotes"
Sub UpdateQuotes()

    Application.StatusBar = "Updating Quotes"
    ActiveWorkbook.Connections("Query - Quotes").Refresh
    Application.StatusBar = ""
End Sub
