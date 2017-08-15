Private Function GetFileFromPicker(Optional strPath As String)
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant

   ' Set up the File Dialog. '
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog
      'Set Inital Path
      If strPath = vbNullString Then
        strPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\"
      End If
    
      .InitialFileName = strPath
      
      ' Allow user to make multiple selections in dialog box '
      .AllowMultiSelect = False

      ' Set the title of the dialog box. '
      .title = "Please select file to import"

      ' Clear out the current filters, and add our own.'
      .Filters.Clear
      .Filters.Add "Excel Spreadsheets", "*.xls;*.xlsx"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the '
      ' user picked at least one file. If the .Show method returns '
      ' False, the user clicked Cancel. '
      If .Show = True Then

         'Loop through each file selected and add it to our list box. '
         GetFileFromPicker = .SelectedItems(1)
      Else
         GetFileFromPicker = ""
      End If
   End With
End Function
