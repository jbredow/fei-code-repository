Option Explicit

Sub reset_application()

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    With ActiveSheet
        .UsedRange
    End With
        
End Sub
