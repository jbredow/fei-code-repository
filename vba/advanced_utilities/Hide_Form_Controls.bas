Private Sub Hide_Form_Controls(xlSh As Worksheet)
    'Remove shape buttons that call VBA code
    Dim shp As shape
    Dim testStr As String

    For Each shp In xlSh.Shapes
        If shp.name Like "VBA*" Then
            shp.Delete
            'shp.Visible = False
        End If
    Next shp
End Sub
