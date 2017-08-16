Sub copy_paste_values()
    Selection.Copy
    Selection.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Application.CutCopyMode = False
End Sub
