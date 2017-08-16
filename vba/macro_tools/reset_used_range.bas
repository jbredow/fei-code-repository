Option Explicit

Sub reset_used_range()
    With ActiveSheet
        Debug.Print .UsedRange.Address(0, 0)
        '.UsedRange.Clear
        .UsedRange    '<~~ called by itself will reset it
        Debug.Print .UsedRange.Address(0, 0)
    End With

End Sub
