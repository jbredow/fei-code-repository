Private Function GetFileExtension(strSave As String)
    Dim strFileExt As String
    
    strFileExt = ""
    If InStrRev(strSave, ".") <> 0 Then
        strFileExt = Right$(strSave, Len(strSave) - InStrRev(strSave, ".") + 1)
    End If

    GetFileExtension = strFileExt
End Function
