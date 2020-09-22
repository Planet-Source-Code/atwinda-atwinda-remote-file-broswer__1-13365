Attribute VB_Name = "modUserPass"
Public ChangedUsers As String
Public ChangedPermissions As String
Public ChangedLogPath As String

Function SaveList(lb As ListBox, Path As String) 'saves a listbox to text file
Open Path For Output As #1 'open the file first, other wide, data will write over the other data about to be entered
For i = 0 To lb.ListCount
    Print #1, lb.List(i)
Next i
Close #1
End Function

Function ReadFile(Path As String, lb As ListBox) 'read text and put it in a listbox, unless it's blank
    Dim fnum As Integer
    Dim sTemp As String
    fnum = FreeFile()
    lb.Clear
    Open Path For Input As #1
    While Not EOF(1)
        Line Input #1, sTemp
        If sTemp = "" Then
        Else
        Call AddItem(sTemp, lb) 'this other wise it still adds it
        End If
   Wend
    Close #1
End Function

Function AddItem(Info As String, lb As ListBox) 'so simple...
lb.AddItem Info
End Function

'I found this on PSC, i don't remember when, and i don't remember the name. Sorry:(
Function encryptAll(data As String, seed As Long) As String
    Dim X As Integer, tmp As String, stepnum As Integer
    Dim byteArray() As Byte, seedOffset As Integer, n As String
    tmp = Trim$(Str(seed))
    seed = 0


    For X = 1 To Len(tmp)
        n = Mid(tmp, X, 1)
        seed = seed + CLng(n)
    Next X
    
reCheckSeed:


    If seed > 255 Then
        seed = -1 + (seed - 255)
        GoTo reCheckSeed
    End If


    For X = 1 To Len(data)
        ReDim Preserve byteArray(X)
        n = Mid(data, X, 1)
        byteArray(X) = Asc(n)
        
        stepnum = seed + X
reCheckstepnum:


        If stepnum > 255 Then
            stepnum = -1 + (stepnum - 255)
            GoTo reCheckstepnum
        End If
        
        byteArray(X) = byteArray(X) Xor CByte(stepnum)
        
    Next X
    tmp = ""

For X = 1 To Len(data)
        tmp = tmp & Chr(byteArray(X))
    Next X
    encryptAll = tmp
End Function



