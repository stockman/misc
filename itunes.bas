Sub maker()
On Error GoTo fin
    'Call fresher
    Application.ScreenUpdating = False
    a = "$C$15"
    dels = 0
    'find locations
    'seachfor = "<key>Location<"
    seachfor = "<key>Rating<"
borjan:
    If ActiveCell.Row = 120000 Then
        GoTo fin
    End If
    Range(Range(a).Offset(0, 1).Address, "$D$150000").Select
    Selection.Find(What:=seachfor, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    sir = ActiveCell.Address
    
    'get relative start and end points
    Cells(ActiveCell.Row, 3).Select
    Selection.End(xlDown).Select
        b = ActiveCell.Address
    Selection.End(xlUp).Select
    ActiveCell.Offset(-1, 0).Select
    a = ActiveCell.Address
    'fix it
    Range(sir).Select
    Call DoStuff(a, b, dels)
    
    GoTo borjan
fin:
    Application.ScreenUpdating = True
    Range("C1").Select
    MsgBox "you deleted " & dels & " songs"
End Sub
Sub searcherton(a, dels, searchfor)
    'initalize
    a = "$C$15"
    seachfor = "<key>Location<"
    Call roll_it(searchfor, a)
    
    seachfor = "<key>Rating<"
End Sub

Sub DoStuff(a, b, dels)
    'mac files
    Base_Files = "file://localhost/Volumes/Backup/T%20Drive/Musik"
    'pac files
    'base files = "file://localhost/"
    filet = ActiveCell.Value
    
    filet = Mid(filet, InStr(1, filet, "<string>") + 8, 9999)
    filet = Replace(filet, Base_Files, "G:\T Drive\Musik")
    filet = Replace(filet, "%20", " ")
    filet = Replace(filet, "</string>", "")
    filet = Replace(filet, "/", "\")
    ActiveCell.Offset(0, 1).Value = filet
    ActiveCell.Offset(0, 2).FormulaR1C1 = _
        "=INDEX(Files!R3C1:R55000C1,MATCH(RC[-1],(Files!R3C1:R55000C1),0),1)"
    resp = ActiveCell.Offset(0, 2).Value
        If IsError(resp) Then
            'delete the xml.
            dels = dels + 1
            Range(a, b).EntireRow.Select
            Selection.Delete Shift:=xlUp
            Cells(ActiveCell.Row, 3).Select
            Else
            ActiveCell.Offset(0, 2).Value = ""
            ActiveCell.Offset(0, 1).Value = ""
            Selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(1, -1).Select
        End If

    a = ActiveCell.Address
End Sub
