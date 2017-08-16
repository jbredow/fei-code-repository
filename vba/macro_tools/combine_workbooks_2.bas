Option Explicit

Sub combine_workbooks_2()
    
    ' consolidates all workbooks from a single folder into a single worksheet
    
    Dim MyPath As String, FilesInPath As String
    Dim MyFiles() As String
    Dim SourceRcount As Long, Fnum As Long
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim sourceRange As Range, destRange As Range
    Dim rNum As Long, CalcMode As Long
    Dim lastRow As Long, LastCol As Long
    
'   Change this to the path\folder location of your files.
    MyPath = "c:\dl\combine"
    
'   Add a slash at the end of the path if needed.
    If Right(MyPath, 1) <> "\" Then
        MyPath = MyPath & "\"
    End If
    
'   If there are no Excel files in the folder, exit.
    FilesInPath = Dir(MyPath & "*.csv*")
    If FilesInPath = "" Then
        MsgBox "No files found"
        Exit Sub
    End If
    
'   Fill the myFiles array with the list of Excel files in the search folder.
    Fnum = 0
    Do While FilesInPath <> ""
        Fnum = Fnum + 1
        ReDim Preserve MyFiles(1 To Fnum)
            MyFiles(Fnum) = FilesInPath
            FilesInPath = Dir()
        Loop
        
    '   Set various application properties.
        With Application
            CalcMode = .Calculation
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
        End With
        
        ' Add a new workbook with one sheet.
        Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
        rNum = 1
        
        ' Loop through all files in the myFiles array.
        If Fnum > 0 Then
            For Fnum = LBound(MyFiles) To UBound(MyFiles)
                Set mybook = Nothing
                On Error Resume Next
                Set mybook = Workbooks.Open(MyPath & MyFiles(Fnum))
                On Error GoTo 0
                
                If Not mybook Is Nothing Then
        On Error Resume Next
        
        ' Change this range to fit your own needs.
        With mybook.Worksheets(1)
            With ActiveSheet
                lastRow = .Range("A" & Rows.Count).End(xlUp).Row
                LastCol = .Range("IV1").End(xlToLeft).Column
            End With
            Set sourceRange = .Range(Cells(1, 1), Cells(lastRow, LastCol))
        End With
            
        If Err.Number > 0 Then
            Err.Clear
            Set sourceRange = Nothing
            Else
                ' If source range uses all columns then skip this file.
            If sourceRange.Columns.Count >= BaseWks.Columns.Count Then
                Set sourceRange = Nothing
            End If
        End If
        On Error GoTo 0
        
        If Not sourceRange Is Nothing Then
        
            SourceRcount = sourceRange.Rows.Count
            
            If rNum + SourceRcount >= BaseWks.Rows.Count Then
                MsgBox "There are not enough rows in the target worksheet."
                BaseWks.Columns.AutoFit
                mybook.Close savechanges:=False
                GoTo ExitTheSub
            Else
            
            ' Copy the file name in column A.
                With sourceRange
                    BaseWks.Cells(rNum, "A"). _
                    Resize(.Rows.Count).Value = MyFiles(Fnum)
                End With
                
                ' Set the destination range.
                Set destRange = BaseWks.Range("B" & rNum)
            
            ' Copy the values from the source range _
                to the destination range.
                With sourceRange
                    Set destRange = destRange. _
                    Resize(.Rows.Count, .Columns.Count)
                End With
                destRange.Value = sourceRange.Value
                
                rNum = rNum + SourceRcount
                End If
            End If
            mybook.Close savechanges:=False
        End If
        
    Next Fnum
    BaseWks.Columns.AutoFit
    End If
    
ExitTheSub:
    ' Restore the application properties.
    With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = CalcMode
    End With
End Sub
