Attribute VB_Name = "modListbox"
Function CheckList(lb As ListBox, sString As String) As Boolean ' this checks for a string in listbox
For i = 0 To lb.ListCount
        If lb.List(i) = sString Then
        CheckList = True: GoTo done
        End If
    Next i
    CheckList = False
done:
End Function

