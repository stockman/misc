Attribute VB_Name = "Module1"
Sub fresher()
  Sheets("XML Worked").Select
    Cells.Select
    Selection.Delete Shift:=xlUp

    Sheets("XML raw").Select
    Cells.Select
    Selection.Copy
    Sheets("XML Worked").Select
    Range("a1").Select
    ActiveSheet.Paste
End Sub
Sub doer()
    'Call fresher
    Sheets("XML Worked").Select
    'get range
    Range("C1").Select
    Selection.End(xlDown).Select
nexter:
    If ActiveCell.Row > 5000 Then
    GoTo fin
    End If
    
    a = ActiveCell.Address
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    b = ActiveCell.Address
    ' delete if missing

    Range(a).Select
borjan:
    If a = b Then
    GoTo nexter
    End If
    tt = ActiveCell.Offset(0, 1).Value
    'check ratings
    Rate = InStr(1, tt, "Rating")
        If Rate <> 0 Then
            Call rater(a, b, tt)
            GoTo nexter
        End If
    'check locations
    loca = InStr(1, tt, "Location")
        If loca <> 0 Then
            Call fileter(a, b, tt)
            GoTo nexter
        End If
        
    ActiveCell.Offset(1, 0).Select
    GoTo borjan
    
fin:
End Sub
Sub fileter(a, b, tt)
    filet = Replace(tt, "%20", " ")
    filet = Replace(filet, "file://localhost/", "")
    filet = Mid(filet, InStr(1, filet, "<string>") + 8, 999)
    filet = Replace(filet, "</string>", "")
    filet = Replace(filet, "/", "\")
    ActiveCell.Offset(0, 2).Value = filet
    ActiveCell.Offset(0, 3).FormulaR1C1 = _
        "=INDEX(Files!R3C1:R55000C1,MATCH(RC[-1],(Files!R3C1:R55000C1),0),1)"
    resp = ActiveCell.Offset(0, 3).Value
        If IsError(resp) Then
            'delete the xml.
            Range(a, b).EntireRow.Select
            Selection.Delete Shift:=xlUp
            Cells(ActiveCell.Row, 3).Select
            Else
            ActiveCell.Offset(0, 3).Value = ""
            ActiveCell.Offset(0, 2).Value = ""
            Selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
        End If
End Sub

Sub rater(a, b, tt)
'do something here
End Sub

