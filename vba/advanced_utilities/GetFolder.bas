Private Function GetFolder(strPath As String)
    If strPath = "" Then
        strPath = "My Documents"
    End If
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = strPath
        .title = "Select Folder"
        If .Show = True Then
            GetFolder = .SelectedItems(1)
        Else
            GetFolder = ""
        End If
    End With

End Function
