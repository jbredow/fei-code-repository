Private Sub GetNames_Form_Controls()
'Dave Peterson and Bob Phillips
'Example only for the Forms controls
    Dim shp As shape
    Dim testStr As String
    Dim xlSh As Worksheet
    
    Set xlSh = ActiveSheet
    
    For Each shp In xlSh.Shapes
        If shp.name Like "VBA*" Then
           'shp.Delete
            shp.Visible = False
            shp.Visible = True
        End If
    Next shp
End Sub
