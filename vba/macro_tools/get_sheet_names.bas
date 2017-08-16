Sub get_sheet_names()
    'Lists each sheet starting with active cell
    Dim wSheet As Worksheet
    
    For Each wSheet In Worksheets
        On Error Resume Next
        ActiveCell.Value = wSheet.name
        ActiveCell.Offset(1, 0).Select
    Next wSheet

End Sub
