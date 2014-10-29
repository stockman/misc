VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
Application.ScreenUpdating = False
Select Case Target.EntireColumn.Address
    Case "$C:$C"
    hello = ActiveCell.Address

    Rows("1:999").Select
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort.SortFields.Add Key:=Range("C2:C999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort
        .SetRange Rows("1:999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range(hello).Select
End Select
Application.ScreenUpdating = True
End Sub

