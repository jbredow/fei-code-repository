Private Function GetExportGroupFileName(strPath As String)
    Dim v As Variant, saveName As String
    
    saveName = CStr(Range("Export_Group_Name").Value)
    If saveName = "" Then saveName = "Weekly Summary Report"
    
    saveName = BuildNewDatedName("", saveName)
    
    v = Application.GetSaveAsFilename(strPath & saveName, "Excel Files (*.xlsx), *.xlsx")
    
    If VarType(v) = vbString Then
        GetExportGroupFileName = v
      Else
        GetExportGroupFileName = ""
    End If
    
End Function
