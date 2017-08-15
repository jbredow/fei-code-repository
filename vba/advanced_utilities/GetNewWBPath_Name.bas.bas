Private Function GetNewWBPath_Name(getName As Boolean) As String
    Dim filePath As String, saveName As String
    
    filePath = CStr(Range("Export_Path").Value)
    
    If getName Then
        'Return full path and filename
        filePath = GetExportGroupFileName(filePath)
        filePath = GetUniqueFileName(filePath)
      Else
        'Only return a path
        filePath = GetFolder(filePath)
        
        If filePath <> "" Then
            'Save Path to Lookups sheet
            If Right(filePath, 1) <> "\" Then filePath = filePath & "\"
            Range("Export_Path").Value = filePath
        End If
    End If
    
    If filePath = "" Then
        Exit Function
    End If
    
    GetNewWBPath_Name = filePath
End Function
