Private Function Test_File_Exists(filePath As String) As Boolean
    Dim fso As Scripting.FileSystemObject

    If filePath = "" Then
        Test_File_Exists = False
        Exit Function
    End If
    
    Set fso = New Scripting.FileSystemObject

    Test_File_Exists = fso.FileExists(filePath)

End Function
