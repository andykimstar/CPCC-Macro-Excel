Sub ScrollBottomButton()
' Last Edit: 2025-01-02

    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet   ' Refer to your sheet instead of ActiveSheet.
    
    targetSheet.Activate    ' Make sure the worksheet is active to prevent errors.
    
    With targetSheet.Cells(targetSheet.Rows.Count, Selection.Column).End(xlUp)
        .Select ' not required to change the focus/view
        ActiveWindow.ScrollRow = .row
    End With
    
End Sub


