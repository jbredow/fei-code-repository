Private Sub HideReportOptions()
    Dim xlApp As Application
    Dim xlWB As Workbook
    Dim xlSh As Worksheet
    
    Set xlApp = Application
    Set xlWB = xlApp.ActiveWorkbook

    Set xlSh = xlWB.Sheets("Report")
    
    xlSh.Shapes("VBA Options").Visible = Not xlSh.Shapes("VBA Options").Visible

End Sub
