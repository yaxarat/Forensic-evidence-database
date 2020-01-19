Sub searchData()

Dim i As Long, lastRow As Long

lastRow = Cells.Find(What:="*", _
After:=Range("A1"), _
LookAt:=xlPart, _
LookIn:=xlFormulas, _
SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious, _
MatchCase:=False).Row

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

For i = 39 To lastRow
    
        If Cells(i, 8) = Range("C9") And Cells(i, 10) = Range("C11") And Cells(i, 14) = Range("C15") Then
            Range("C5") = Cells(i, 2)
            Range("C6") = Cells(i, 3)
            Range("C7") = Cells(i, 4)
            Range("C8") = Cells(i, 5)
            Range("C9") = Cells(i, 6)
            Range("C10") = Cells(i, 7)
            Range("C11") = Cells(i, 8)
            Range("C12") = Cells(i, 9)
            Range("C13") = Cells(i, 10)
            Range("C14") = Cells(i, 11)
            Range("C15") = Cells(i, 12)
            Range("C16") = Cells(i, 13)
            Range("C17") = Cells(i, 14)
            Range("C18") = Cells(i, 15)
            Range("C19") = Cells(i, 16)
            Range("C20") = Cells(i, 17)
            Range("C21") = Cells(i, 18)
            Range("C22") = Cells(i, 19)
            Range("C23") = Cells(i, 20)
            Range("C24") = Cells(i, 21)
            Range("C25") = Cells(i, 22)
            Range("C26") = Cells(i, 23)
            Range("C27") = Cells(i, 24)
            Range("C28") = Cells(i, 25)
            Range("C29") = Cells(i, 26)
            Range("C30") = Cells(i, 27)
            Range("C31") = Cells(i, 28)
            Range("C32") = Cells(i, 29)
            Range("C33") = Cells(i, 30)
            Range("C34") = Cells(i, 31)
            Range("C35") = Cells(i, 32)
            Range("C36") = Cells(i, 33)
            Cells(16, 10) = i
        Exit Sub
        End If
Next i

MsgBox lastRow

MsgBox "Record doesn't exist"
        
End Sub