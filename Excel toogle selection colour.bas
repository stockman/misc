Attribute VB_Name = "Module2"
Sub tempcolour()
'toggles the backround of a selection's rows so that you can easily visually inspect
Dim addr As String
    addr = Selection.Address

If Selection.Interior.Color = 16771499 Then
    Selection.EntireRow.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Else
    Selection.EntireRow.Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16771499
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

Range(addr).Select

End Sub
