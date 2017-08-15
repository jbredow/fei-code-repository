Private Sub CleanWBCopy(xlWB As Workbook)
    'Remove any form items from the WB that might require code
    Hide_Form_Controls xlWB.Sheets("Report")
    Hide_Form_Controls xlWB.Sheets("Data")
    Hide_Form_Controls xlWB.Sheets("Resources")
    
    'Update the data source, otherwise points to original WB
    xlWB.Sheets("Report").PivotTables("Pivot_Report").SourceData = "Table_Data"
    xlWB.ShowPivotTableFieldList = False
    
    'Remove SQL Data Connections
    xlWB.Connections("a02963 ReportData").Delete
    xlWB.Connections("a02953 PCList").Delete
End Sub
