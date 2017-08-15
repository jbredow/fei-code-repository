Private Sub SaveAsNewWB_NoCode(saveName As String)
    Dim xlWB As Workbook
    
    'Create new WB without module
    Set xlWB = SaveWBWithoutCode
    CleanWBCopy xlWB
    
    Application.DisplayAlerts = False
        xlWB.SaveAs saveName, xlWorkbookDefault
        xlWB.Close
    Application.DisplayAlerts = True
    
    Set xlWB = Nothing
End Sub
