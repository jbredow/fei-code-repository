Sub firstColumnTabSplitter()
    ' FirstColumnTabSplitter Macro / assumes column "A" as unique list
    ' Sort by Column A and will make tabs based on unique values
    ' Updated 11.27.12 to work for both .xlsx and .xls

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim rRange          As Range
    Dim rCell           As Range
    Dim wSheet          As Worksheet
    Dim wSheetStart     As Worksheet
    Dim strText         As String
    Dim lastRow         As Long
 
    Set wSheetStart = ActiveSheet
    wSheetStart.AutoFilterMode = False

    'Set a range  variable to the correct item column
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set rRange = Range("A1:A" & lastRow)

    'Delete any sheet called "UniqueList"
    'Turn off run time  errors &  delete alert
    On Error Resume Next
    Worksheets("UniqueList").Delete
    'Add a sheet called "UniqueList"
    Worksheets.Add().name = "UniqueList"

    'Filter the Set range so only a unique list is created
    With Worksheets("UniqueList")
        rRange.AdvancedFilter xlFilterCopy, , _
        Worksheets("UniqueList").Range("A1"), True

        'Set a range variable to the unique list, less the  heading.
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Set rRange = Range("A2:A" & lastRow)
    End With

    On Error Resume Next
    With wSheetStart
        For Each rCell In rRange
            strText = rCell
            .Range("A1").AutoFilter 1, strText
            Worksheets(strText).Delete
             'Add a sheet named as content of rCell
            Worksheets.Add().name = strText
             'Copy the visible filtered range _

'            (default of Copy Method) And leave hidden rows

            .UsedRange.Copy Destination:=ActiveSheet.Range("A1")
            ActiveSheet.Cells.Columns.AutoFit
        Next rCell
    End With

    With wSheetStart
        .AutoFilterMode = False
        .Activate
    End With

    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
