Private Function GetUniqueFileName(strSave As String)
    Dim strTest As String, strFileExt As String, x As Integer
    
    'Create a new unique filename if file exists
    If Test_File_Exists(strSave) Then
        strFileExt = GetFileExtension(strSave)
        
        x = 1
        Do
            strTest = Replace(strSave, strFileExt, "(" & CStr(x) & ")" & strFileExt)
            x = x + 1
        Loop While Test_File_Exists(strTest)
        strSave = strTest
    End If

    GetUniqueFileName = strSave
End Function
