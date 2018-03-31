Attribute VB_Name = "Helpers"
Sub DisableApplication()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
End Sub

Sub EnableApplication()
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
