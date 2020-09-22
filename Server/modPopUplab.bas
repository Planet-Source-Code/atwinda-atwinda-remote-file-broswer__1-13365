Attribute VB_Name = "modPopUplab"
Dim Message As String
Dim MessType As String
Dim Title As String

Function ParseMessagepart1(Info As String) 'breaks the message off from the type and title
Dim strTypeandTitle As String

Message = Left$(Info$, InStr(Info$, "|") - 1)
strTypeandTitle$ = Right$(Info$, Len(Info$) - InStr(Info$, "|"))

Call ParseMessagepart2(strTypeandTitle)
End Function

Function ParseMessagepart2(Info As String) ' breaks the type from title
MessType = Left$(Info$, InStr(Info$, "|") - 1)
Title = Right$(Info$, Len(Info$) - InStr(Info$, "|"))

Call PopUpMessage
End Function

Function PopUpMessage() 'selects the type of popup, and pops it up!
Select Case MessType
    Case "vbCritical":
        Call MsgBox(Message$, vbCritical, Title$)
    Case "vbExclamation":
        Call MsgBox(Message$, vbExclamation, Title$)
    Case "vbInformation":
        Call MsgBox(Message$, vbInformation, Title$)
    Case "vbQuestion":
        Call MsgBox(Message$, vbQuestion, Title$)
    Case "vbOKOnly":
        Call MsgBox(Message$, vbOKOnly, Title$)
    Case "vbSystemModal":
        Call MsgBox(Message$, vbSystemModal, Title$)
End Select
End Function
