Attribute VB_Name = "modCommand"
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_SHUTDOWN As Long = 1

Public Connected As String
Public strProcess As String
Public Logedin As String
Public Logging As String
Public LoggingPath As String
Public Permissions As String
Dim strList As String
'this is where most of the fun stuff is...

Function DataParsing(theData As String, wsk As Winsock) 'cuts the incoming data into readable bits.
On Error Resume Next
Dim Command As String, Info As String

Command$ = Left$(theData$, InStr(theData$, "|") - 1)
Info$ = Right$(theData$, Len(theData$) - InStr(theData$, "|"))

Select Case Command

Case "UserPass": 'if the client sent a username and PW
    Call AddText("Checking username and password.")
    Call CheckUserPass(Info)

Case "Dir": 'client requests a directory list
    If Permissions = "Yes" Then 'check for permissions
        If CheckList(frmMain.lst4, "Dir") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then ' all commands has this, it just makes sure that someone is loggedin
    Call AddText("Requested directory list: " & Info)
    Call DirList(Info, frmMain.Dir, wsk) 'kick it to the dirlist function
    End If

Case "File": 'file request
        If Permissions = "Yes" Then 'check permissions
        If CheckList(frmMain.lst4, "File") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Requested file list: " & Info)
    Call FileList(Info, frmMain.File, wsk) 'kick to the filelist command
    End If
    
Case "SaveFile": 'the savefile command (used in file editing, and file creating)
     If Permissions = "Yes" Then 'again, permissions
        If CheckList(frmMain.lst4, "SaveFile") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Requested to save file.")
    Call SaveFile(Info, wsk) 'what do u think this does?
    End If
    
Case "ViewFile": 'the view file command...
     If Permissions = "Yes" Then
        If CheckList(frmMain.lst4, "ViewFile") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Requested to view file.")
    Call ViewFile(Info, wsk) 'gee... do i even need to say? yes, ok. it kicks it to the viewfile function
    End If
    
Case "EditFile": 'the first part of file editing, is the same as viewing.
     If Permissions = "Yes" Then
        If CheckList(frmMain.lst4, "EditFile") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Requested to edit a file.")
    Call EditFile(Info, wsk) 'wow... the edit file function...
    End If
    
Case "PopUp": 'the popup 'lab' is in another mod, there was just to many things to put in this one...
     If Permissions = "Yes" Then
        If CheckList(frmMain.lst4, "PopUp") = True Then
        wsk.SendData ("Perm")
        Exit Function
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Recivied PopUp Message.")
    Call ParseMessagepart1(Info) 'kick it to the other mod
    End If

Case Else: Call AddText("Error: Bad Command or Filename") 'if it's anything else, then send an error to the client
    Call wsk.SendData("Error")
End Select
End Function

Function AddText(Text As String) 'a one command does it all debug function (also writes to the log if thats enabled)
frmMain.txtDisplay.Text = frmMain.txtDisplay.Text & Text & vbCrLf
frmMain.txtDisplay.SelStart = Len(frmMain.txtDisplay.Text)
frmMain.txtDisplay.SelLength = 0
If Logging = "Yes" Then 'if the log is on, then write it
Call Log(LoggingPath, Text & " " & Time)
End If
End Function

Function DriveList(Drive As DriveListBox, wsk As Winsock) 'gets all the contents from a drivelistbox
On Error Resume Next
strList = "" 'clears the temp list
For i = 0 To Drive.ListCount 'breaks the drivelist down
    strList = strList & Drive.List(i) & vbCrLf
Next i
Call wsk.SendData("DriveList" & "|" & strList) 'and sends it!
Call AddText("Filled drive list request.") 'adds text to the window
End Function

Function DirList(DirPath As String, Dir As DirListBox, wsk As Winsock) 'basically the same as the drivelist, but with alittle error handling
strList = ""
On Error GoTo CheckErr 'there it is, the 'error handling'
Dir.Path = DirPath
DoEvents
For i = 0 To Dir.ListCount
    strList = strList & Dir.List(i) & vbCrLf
Next i
DoEvents
Call wsk.SendData("DirList" & "|" & strList)
Call AddText("Filled directory list request: " & DirPath)

CheckErr: 'i got 53 and 73 errors when i made it mess up(aka non-valid path) so thats that
If Err = 53 Then
    Call SendNoPath(DirPath, wsk)
   Exit Function
ElseIf Err = 73 Then
     Call SendNoPath(DirPath, wsk)
    Exit Function
Else: Exit Function 'just incase anything else goes wrong
End If
End Function

Function FileList(DirPath As String, File As FileListBox, wsk As Winsock) 'exact same, but with a filelistbox...
strList = ""
On Error GoTo CheckErr
File.Path = DirPath
DoEvents
For i = 0 To File.ListCount
    strList = strList & File.List(i) & vbCrLf
Next i
DoEvents
Call wsk.SendData("FileList" & "|" & strList)
Call AddText("Filled file list request: " & DirPath)

CheckErr:
If Err = 53 Then
   Call SendNoPath(DirPath, wsk)
   Exit Function
ElseIf Err = 73 Then
     Call SendNoPath(DirPath, wsk)
    Exit Function
Else: Exit Function
End If
End Function

Function SendNoPath(Path As String, wsk As Winsock) 'if non-valid path, send it back to the client
On Error Resume Next
If Connected = "Yes" Then
Call wsk.SendData("NoPath" & "|" & Path)
End If
End Function

Function Shutdown() 'the shoutdown command... look out...
Dim lngresult
Call frmMain.wskFile.SendData("ServerDown")
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Function

Function ListProcess(wsk As Winsock) 'call the list process command in the other mod (ModPocess)
If Connected = "Yes" Then
    Call FillWindows(wsk)
End If

End Function

Function SaveFile(Text As String, wsk As Winsock) 'the savefile comes in one piece, and needs to be broken down
On Error Resume Next
Dim strLocation As String, strText As String

strLocation = Left$(Text$, InStr(Text$, "š") - 1) 'the path to save to...
strText = Right$(Text$, Len(Text$) - InStr(Text$, "š")) 'the teaxt to save to the path...

Open strLocation$ For Output As #1 'u r standard open for output
Print #1, strText$
Close #1
Call wsk.SendData("MadeFile" & "|" & Location) 'tell the client the file was made/edited
End Function

Function ViewFile(Path As String, wsk As Winsock) 'open the whole file, and sends it to the client
On Error Resume Next
If CheckPath(Path) = True Then
Dim FileContents As String
Dim NumFile As Long
NumFile = FileLen(Path)
Open (Path) For Input As #1
FileContents = Input(NumFile, #1)
Close #1
DoEvents
Call wsk.SendData("ViewFile" & "|" & FileContents)
NumFile = 0
FileContents = ""
Else: wsk.SendData ("NoPath" & "|" & Path)
End If
End Function

Function EditFile(Path As String, wsk As Winsock) 'same as view file, but the client need to know that it
'needs to setup the textlab alittle differently
On Error Resume Next
If CheckPath(Path) = True Then
Dim FileContents As String
Dim NumFile As Long
NumFile = FileLen(Path)
Open (Path) For Input As #1
FileContents = Input(NumFile, #1)
Close #1
DoEvents
Call wsk.SendData("EditFile" & "|" & Path & "|" & FileContents)
NumFile = 0
FileContents = ""
Else: wsk.SendData ("NoPath" & "|" & Path)
End If
End Function

Function Log(Path As String, Info As String) 'the logging function... wow...
On Error Resume Next
Open Path$ For Append As #1
Print #1, Info$
Close #1
End Function

Function CheckUserPass(Info As String) 'if userpass is required, and the client sends that data
On Error Resume Next
Dim User As String, Pass As String, EnPass As String

User$ = Left$(Info$, InStr(Info$, "|") - 1) ' the username
Pass$ = Right$(Info$, Len(Info$) - InStr(Info$, "|")) 'the password

If CheckList(frmMain.lst1, User$) = True Then
    EnPass = encryptAll(Pass$, "219119") 'encrypts the password to check it
    If CheckList(frmMain.lst2, EnPass) = True Then
    Logedin = "Yes" 'set logedin to yes
    Call AddText("User: " & User$ & " Logged in.") 'send it to the client
    Call TellLogedIn("Yes")
    Else: Call TellLogedIn("No") 'bad login..
    Exit Function
    End If
Else: Call TellLogedIn("No")
End If

End Function

Function TellLogedIn(YesNo As String) 'send to the client that there logged in...
On Error Resume Next
If YesNo = "Yes" Then
Call frmMain.wskFile.SendData("Login" & "|" & "Yes")
ElseIf YesNo = "No" Then
Call frmMain.wskFile.SendData("Login" & "|" & "No") 'or they aren't
End If
End Function

Function CheckPath(FileName As String) As Boolean 'got this from PSC, don't remeber the name, sorry :(
    Dim temp As String
    On Error GoTo NotFound
    temp = Dir(FileName)
    If temp = "" Then CheckPath = False Else CheckPath = True
    temp = ""
NotFound:
    If Err = 53 Then Resume Next
End Function

