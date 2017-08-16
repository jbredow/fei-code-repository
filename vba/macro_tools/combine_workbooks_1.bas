Option Explicit

'32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal _
    pszpath As String) As Long

Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) _
    As Long

Public Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Function GetDirectory(Optional Msg) As String

    On Error Resume Next
    Dim bInfo As BrowseInfo
    Dim Path As String
    Dim r As Long, x As Long, pos As Integer
    
    'Root folder = Desktop
    bInfo.pIDLRoot = 0&
    
    'Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Please select the folder of the excel files to copy."
    Else
        bInfo.lpszTitle = Msg
    End If
    
    'Type of directory to return
    bInfo.ulFlags = &H1
    
    'Display the dialog
    x = SHBrowseForFolder(bInfo)
    
    'Parse the result
    Path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal Path)
    If r Then
        pos = InStr(Path, Chr$(0))
        GetDirectory = Left(Path, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function

Sub combine_workbooks_1()
'   uses above function
    Dim Path As String
    Dim filename As String
    Dim LastCell As Range
    Dim Wkb As Workbook
    Dim ws As Worksheet
    Dim ThisWB As String
    
    ThisWB = ThisWorkbook.name
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Path = GetDirectory
    filename = Dir(Path & "\*.csv", vbNormal) 'filter
    Do Until filename = ""
        If filename <> ThisWB Then
            Set Wkb = Workbooks.Open(filename:=Path & "\" & filename)
            For Each ws In Wkb.Worksheets
                Set LastCell = ws.Cells.SpecialCells(xlCellTypeLastCell)
                If LastCell.Value = "" And LastCell.Address = Range("$A$1").Address Then
                Else
                    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                End If
            Next ws
            Wkb.Close False
        End If
        filename = Dir()
    Loop
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Set Wkb = Nothing
    Set LastCell = Nothing
End Sub
