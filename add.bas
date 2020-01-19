Attribute VB_Name = "Module1"
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

If Range("H19") = "" Then
    MsgBox "You didn't enter the diameter of circle!"
    Range("H19").Select
    Exit Sub
    End If

If Range("H15") = "" Then
    MsgBox "You didn't enter the thickness!"
    Range("H15").Select
    Exit Sub
    End If

If Range("H13") = "" Then
    MsgBox "You didn't enter the overall width!"
    Range("H13").Select
    Exit Sub
    End If

Cells(nextBlankRow, 9) = Range("H9")
Cells(nextBlankRow, 10) = Range("H10")
Cells(nextBlankRow, 11) = Range("H11")
Cells(nextBlankRow, 12) = Range("H12")
Cells(nextBlankRow, 13) = Range("H13")
Cells(nextBlankRow, 14) = Range("H14")
Cells(nextBlankRow, 15) = Range("H15")
Cells(nextBlankRow, 16) = Range("H16")
Cells(nextBlankRow, 17) = Range("H17")
Cells(nextBlankRow, 18) = Range("H18")
Cells(nextBlankRow, 19) = Range("H19")
Cells(nextBlankRow, 20) = Range("H20")
Cells(nextBlankRow, 21) = Range("H21")
Cells(nextBlankRow, 22) = Range("H22")
Cells(nextBlankRow, 23) = Range("H23")
Cells(nextBlankRow, 24) = Range("H24")
Cells(nextBlankRow, 25) = Range("H25")
Cells(nextBlankRow, 26) = Range("H26")
Cells(nextBlankRow, 27) = Range("H27")
Cells(nextBlankRow, 28) = Range("H28")
Cells(nextBlankRow, 29) = Range("H29")
Cells(nextBlankRow, 30) = Range("H30")
Cells(nextBlankRow, 31) = Range("H31")
Cells(nextBlankRow, 32) = Range("H32")
Cells(nextBlankRow, 33) = Range("H33")
Cells(nextBlankRow, 34) = Range("H34")
Cells(nextBlankRow, 35) = Range("H35")
Cells(nextBlankRow, 36) = Range("H36")
Cells(nextBlankRow, 37) = Range("H37")
Cells(nextBlankRow, 38) = Range("H38")
Cells(nextBlankRow, 39) = Range("H39")
Cells(nextBlankRow, 40) = Range("H40")

'MsgBox nextBlankRow
Dim p As Long, q As Long
p = 44
q = p + 1
Do While Cells(p, 13) <> ""
    Do While Cells(q, 13) <> ""
        If Cells(p, 13) = Cells(q, 13) And Cells(p, 15) = Cells(q, 15) And Cells(p, 19) = Cells(q, 19) Then
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
