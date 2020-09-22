Attribute VB_Name = "modCommand"
Public Connected As String

Function DataParsing(theData As String, wsk As Winsock)  'cut the data into usable bits
On Error Resume Next
Dim Command As String, Info As String

Command$ = Left$(theData, InStr(theData, "|") - 1)
Info$ = Right$(theData$, Len(theData$) - InStr(theData$, "|"))

'these all just select the command part of the string, and send it to the coresponding function
Select Case Command
Case "DriveList": Call AddText(Info)
Case "DirList": Call AddText(Info)
Case "FileList": Call AddText(Info)
Case "Processes": Call AddText(Info)
Case "MadeFile": Call AddText("File: " & Info & " was made or edited.")
Case "ViewFile": Call ViewFile(Info)
Case "EditFile": Call EditFile(Info)
Case "NoPath": Call AddText("The path: " & Info & " does not exist.")
Case "Login": Call Login(Info)
Case Else: Call AddText("Server sent bad command.")
End Select
End Function

Function CheckCommand(Command As String, wsk As Winsock) 'check the command to be sent
On Error Resume Next
Dim sndCommand As String, sndInfo As String

sndCommand = Left(Command, InStr(Command, "|") - 1)
sndInfo = Right(Command, Len(Command) - InStr(Command, "|"))

Select Case sndCommand
Case "Dir": Call wsk.SendData("Drive" & "|" & sndInfo)
Case "File": Call wsk.SendData("File" & "|" & sndInfo)
Case Else: Call AddText(" Bad Command or Filename")
End Select
End Function

Function Login(YesNo As String) ' if logged in or not
If YesNo = "Yes" Then
    Call AddText("The login was a succes!")
    frmMain.cmdSend.Enabled = True
    frmMain.txtText.Enabled = True
    frmMain.cmdProclist.Enabled = True
    frmMain.cmdShutdown.Enabled = True
    frmMain.cmdDrive.Enabled = True
ElseIf YesNo = "no" Then
Dim Answer As VbMsgBoxResult
    Answer = MsgBox("The username or password you entered are not valid. please retry.", vbOKCancel, "Error")
        If Answer = vbOK Then
            frmLogin.Show
        ElseIf Answer = vbCancel Then
            frmMain.wskFile.Close
            End
        End If
End If
End Function

Function AddText(Text As String) ' the one function does it all
frmMain.txtDisplay.Text = frmMain.txtDisplay.Text & Text & vbCrLf
frmMain.txtDisplay.SelStart = Len(frmMain.txtDisplay.Text)
frmMain.txtDisplay.SelLength = 0
End Function

Function ViewFile(Text As String) 'open teh textlab for viewing a file(disable save stuff...)
frmTextlab.txtText.Text = Text
frmTextlab.txtText.Locked = True
frmTextlab.mnuSave.Enabled = False
frmTextlab.Show
End Function

Function EditFile(Text As String) 'open the textlab for editing a file(able to save edited)
On Error Resume Next
Dim strPath As String, strText As String

strPath = Left(Text, InStr(Text, "|") - 1)
strText = Right(Text, Len(Text) - InStr(Text, "|"))

frmTextlab.txtText.Text = strText
frmTextlab.txtPath.Text = strPath
frmTextlab.mnuSavenew.Enabled = False
frmTextlab.Show
End Function
