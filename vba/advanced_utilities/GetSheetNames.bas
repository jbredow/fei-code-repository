Private Function GetSheetNames(xlWB As Excel.Workbook)
    'Return a list of sheet names in WB
    Dim xlSh As Excel.Worksheet
    Dim strNames As String
    
    strNames = ""
    
    For Each xlSh In xlWB.Sheets
        If strNames <> "" Then strNames = strNames & "~!~"
        strNames = strNames & xlSh.name
    Next

    GetSheetNames = strNames
End Function
