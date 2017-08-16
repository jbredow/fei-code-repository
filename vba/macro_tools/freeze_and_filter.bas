Sub freeze_and_filter()

    Dim sStartCell As String
    
    sStartCell = ActiveCell.Address

    Application.ScreenUpdating = False
    
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "Accent1"
    
    Range(sStartCell).Select
    
    Application.ScreenUpdating = True
End Sub
