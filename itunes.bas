Sub futurethings()
'check these files
'G:\T Drive\USB card\Key Backed for OSX install\RandWiki\Document Library\Mine\iPhone
'add ratings back into to matching songs?
End Sub

Sub doit()
'process
    '1-collect all song ratings
    '2- delete tracks that have files missing.
On Error GoTo fin
    Call fresher
    Application.ScreenUpdating = False
    a = "$C$15"
    dels = 0
    'things to search
    sircher = "<key>Rating<"
    
borjan:
    If sircher = "done" Then
        GoTo fin
    End If
    Select Case sircher
          Case "<key>Location<"
            Call searcherton(a, b, dels, sircher)
            If sircher <> "done" Then
            Call DoStuff(a, b, dels)
            End If
        Case "<key>Rating<"
            Call searcherton(a, b, dels, sircher)
            If sircher <> "<key>Location<" Then
            Call rater(a, b)
            End If
    Case Else
    End Select
    GoTo borjan
fin:
    Application.ScreenUpdating = True
    Range("C1").Select
    MsgBox "you deleted " & dels & " songs"
End Sub
Sub searcherton(a, b, dels, sircher)
    On Error GoTo UnexpectedError
    Range(Range(a).Offset(0, 1).Address, "$D$150000").Select
    Selection.Find(What:=sircher, After:=ActiveCell, LookIn:=xlFormulas, _
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
    Exit Sub
    
' if find error then switch to location
UnexpectedError:
If sircher = "<key>Rating<" Then
    a = "$C$15"
    sircher = "<key>Location<"
Else
    sircher = "done"
End If
Exit Sub

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

Sub rater(a, b)


    ratedress = ActiveCell.Address
    ratethis = ActiveCell.Value
    Selection.End(xlUp).Select
    rateTrack = ActiveCell.Value
    rateName = ActiveCell.Offset(1, 0).Value
    rateArtist = ActiveCell.Offset(2, 0).Value
    
    Sheets("Ratings").Select
    ActiveCell.Value = ratethis
    ActiveCell.Offset(0, 1).Value = rateTrack
    ActiveCell.Offset(0, 2).Value = rateName
    ActiveCell.Offset(0, 3).Value = rateArtist
    
    ActiveCell.Offset(1, 0).Select
    'reset to next
    Sheets("XML Worked").Select
    Range(b).Select
    ActiveCell.Offset(1, 0).Select
    a = ActiveCell.Address
End Sub
