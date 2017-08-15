Private Function SaveWBWithoutCode() As Workbook
    'Copy all teh sheets of the current WB and return the WB object of the newly created WB
    Dim ws As Worksheet
    Dim i As Integer
    Dim sarrWS() As String
    
    ReDim sarrWS(1 To ThisWorkbook.Worksheets.Count)
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
        sarrWS(i) = ws.name
    Next ws
    
    ThisWorkbook.Worksheets(sarrWS()).Copy
    Set SaveWBWithoutCode = ActiveWorkbook
End Function
