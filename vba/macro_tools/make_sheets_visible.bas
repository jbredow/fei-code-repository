Option Explicit

Sub make_sheets_visible()
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Visible = xlSheetVisible
    Next i
    
    Application.ScreenUpdating = True
    
End Sub
