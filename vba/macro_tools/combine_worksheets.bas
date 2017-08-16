Option Explicit

Sub combine_worksheets()
    Dim j As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For j = 2 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(j).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A1048576").End(xlUp)(2)
    Next
End Sub
