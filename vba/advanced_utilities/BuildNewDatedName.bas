Private Function BuildNewDatedName(filePath As String, filePrefix As String) As String
    'Takes file prefix and appends the selected date range and file extension
    'Then gets a uniquefilename if that file already exists
    Dim DateFrom As String, DateTo As String
    
    DateFrom = Format(Range("Date_From"), "mmddyy")
    DateTo = Format(Range("Date_To"), "mmddyy")
    
    If filePath <> "" Then
        If Right(filePath, 1) <> "\" Then filePath = filePath & "\"
    End If
    
    BuildNewDatedName = GetUniqueFileName(filePath & filePrefix & "_" & DateFrom & "-" & DateTo & ".xlsx")
End Function
