Private Sub testShapes()
    Dim xlWB As Workbook, shape As shape
    Dim numShapes As Integer, x As Integer
    
    Set xlWB = ActiveWorkbook
    
    'xlWB.Sheets("MASTER INPUT SHEET").Shapes(5).Name = "BtnAutoFill"
    numShapes = xlWB.ActiveSheet.Shapes.Count
    Debug.Print numShapes

    For x = 1 To numShapes
        Debug.Print x, xlWB.ActiveSheet.Shapes(x).name, xlWB.ActiveSheet.Shapes(x).TextFrame.Characters.text, xlWB.ActiveSheet.Shapes(x).OnAction
    Next x
End Sub
