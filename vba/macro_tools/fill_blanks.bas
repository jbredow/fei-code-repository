Option Explicit

Sub fill_blanks()
' Fill in the blanks for columns that are only populated with the first
' cell of a particular data type.  Select the area that you want to have
' the cells filled in and run the macro.
    Dim rRange1 As Range, rRange2 As Range
    Dim iReply As Long
    
    If Selection.Cells.Count = 1 Then
        MsgBox "Select list and include the blank cells", _
            vbInformation, "Midwest RPC"
            Exit Sub
    ElseIf Selection.Columns.Count > 1 Then
        MsgBox "You can select only one column", _
            vbInformation, "Midwest RPC"
            Exit Sub
    End If

    Set rRange1 = Range(Selection.Cells(1, 1), _
        Cells(1048576, Selection.Column).End(xlUp))
    On Error Resume Next
    Set rRange2 = rRange1.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If rRange2 Is Nothing Then
        MsgBox "There were NO blank cells Found", _
            vbInformation, "Midwest RPC"
        Exit Sub
    End If
 
    rRange2.FormulaR1C1 = "=R[-1]C"
' uncomment below to allow for formula fill
    'iReply = MsgBox("Convert to Values", vbYesNo + vbQuestion, "Midwest RPC")
    'If iReply = vbYes Then rRange1 = rRange1.value
    rRange1 = rRange1.Value
End Sub
