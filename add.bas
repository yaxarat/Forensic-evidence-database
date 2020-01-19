Sub addData()
Dim i As Long, lastRow As Long, nextBlankRow As Long
lastRow = Cells.Find(What:="*", _
After:=Range("A1"), _
LookAt:=xlPart, _
LookIn:=xlFormulas, _
SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious, _
MatchCase:=False).Row

nextBlankRow = lastRow + 1

If Range("C15") = "" Then
    MsgBox "You didn't enter the diameter of circle!"
    Range("C15").Select
    Exit Sub
    End If

If Range("C11") = "" Then
    MsgBox "You didn't enter the thickness!"
    Range("C11").Select
    Exit Sub
    End If

If Range("C9") = "" Then
    MsgBox "You didn't enter the overall width!"
    Range("C9").Select
    Exit Sub
    End If

Cells(nextBlankRow, 2) = Range("C5")
Cells(nextBlankRow, 3) = Range("C6")
Cells(nextBlankRow, 4) = Range("C7")
Cells(nextBlankRow, 5) = Range("C8")
Cells(nextBlankRow, 6) = Range("C9")
Cells(nextBlankRow, 7) = Range("C10")
Cells(nextBlankRow, 8) = Range("C11")
Cells(nextBlankRow, 9) = Range("C12")
Cells(nextBlankRow, 10) = Range("C13")
Cells(nextBlankRow, 11) = Range("C14")
Cells(nextBlankRow, 12) = Range("C15")
Cells(nextBlankRow, 13) = Range("C16")
Cells(nextBlankRow, 14) = Range("C17")
Cells(nextBlankRow, 15) = Range("C18")
Cells(nextBlankRow, 16) = Range("C19")
Cells(nextBlankRow, 17) = Range("C20")
Cells(nextBlankRow, 18) = Range("C21")
Cells(nextBlankRow, 19) = Range("C22")
Cells(nextBlankRow, 20) = Range("C23")
Cells(nextBlankRow, 21) = Range("C24")
Cells(nextBlankRow, 22) = Range("C25")
Cells(nextBlankRow, 23) = Range("C26")
Cells(nextBlankRow, 24) = Range("C27")
Cells(nextBlankRow, 25) = Range("C28")
Cells(nextBlankRow, 26) = Range("C29")
Cells(nextBlankRow, 27) = Range("C30")
Cells(nextBlankRow, 28) = Range("C31")
Cells(nextBlankRow, 29) = Range("C32")
Cells(nextBlankRow, 30) = Range("C33")
Cells(nextBlankRow, 31) = Range("C34")
Cells(nextBlankRow, 32) = Range("C35")
Cells(nextBlankRow, 33) = Range("C36")

'MsgBox nextBlankRow
Dim p As Long, q As Long
p = 39
q = p + 1
Do While Cells(p, 8) <> ""
    Do While Cells(q, 8) <> ""
        If Cells(p, 8) = Cells(q, 8) And Cells(p, 10) = Cells(q, 10) And Cells(p, 14) = Cells(q, 14) Then
            MsgBox "Duplicate Data! Will be removed from database!"
            Range(Cells(q, 7), Cells(q, 38)).ClearContents
        Else
        q = q + 1
        End If
    Loop
p = p + 1
q = p + 1
Loop

End Sub
