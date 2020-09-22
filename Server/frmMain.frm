VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Remote File: Server"
   ClientHeight    =   3480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskFile 
      Left            =   5160
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lst4 
      Height          =   255
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lst3 
      Height          =   255
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lst2 
      Height          =   255
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lst1 
      Height          =   255
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.DirListBox Dir 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server "
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtDisplay 
         Height          =   2415
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   615
      End
   End
   Begin VB.FileListBox File 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   5400
      System          =   -1  'True
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   5160
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mniOpt 
         Caption         =   "Options"
      End
      Begin VB.Menu spacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnoOther 
      Caption         =   "Other"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Window"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Atwinda Remote File Browsing
'The server side is where all the 'heavy' coding is...

Private Sub cmdListen_Click() 'tells the winsock to listen on port 1289 if the caption in 'listen'
If cmdListen.Caption = "Listen" Then
    wskFile.Close
    wskFile.LocalPort = "1289"
    wskFile.Listen
    lblStatus.Caption = "Waiting for Connection..."
    cmdListen.Caption = "Close"
ElseIf cmdListen.Caption = "Close" Then 'if teh caption is 'close' then close the winsock
    wskFile.Close
    lblStatus.Caption = "Closed..."
    cmdListen.Caption = "Listen"
End If
End Sub

Private Sub Form_Load() 'sets up varables, and checks what functions are in use.
On Error Resume Next
lblStatus.Caption = "Closed..."
Connected = "No"
Logedin = "No"
If ReadINI("UserPass", "Value", App.Path & "\Options.ini") = "Yes" Then
    Call ReadFile(App.Path & "\users.rfd", lst1)
    Call ReadFile(App.Path & "\pass.rfd", lst2)
Else:  Logedin = "Yes"
End If

If ReadINI("Permissions", "Value", App.Path & "\Options.ini") = "Yes" Then
    Call ReadFile(App.Path & "\unblocked.rfd", lst3)
    Call ReadFile(App.Path & "\blocked.rfd", lst4)
    Permissions = "Yes"
Else: Permissions = "No"
End If
If ReadINI("Logging", "Value", App.Path & "\Options.ini") = "Yes" Then
    LoggingPath = ReadINI("Logging", "Path", App.Path & "\Options.ini")
    Logging = "Yes"
Else: Logging = "No"

End If
End Sub

Private Sub Form_Unload(Cancel As Integer) 'closes all windows if the main window in closed
End
End Sub

Private Sub mniOpt_Click() 'shows the options window
frmOptions.Show
End Sub

Private Sub mnuClear_Click() 'clears the display...
txtDisplay.Text = ""
End Sub

Private Sub wskFile_Close() 'if the winsock closes do this stuff
lblStatus.Caption = "Closed..."
Connected = "No"
cmdListen.Caption = "Listen"
Call AddText("Connection Closed.")
End Sub

Private Sub wskFile_ConnectionRequest(ByVal requestID As Long) 'connect to the incoming computer
wskFile.Close
wskFile.Accept requestID

lblStatus.Caption = "Connected..."
Connected = "Yes"
Logedin = "Yes"

'and see if user logon in required

If ReadINI("UserPass", "Value", App.Path & "\Options.ini") = "Yes" Then
Call AddText("Requesting Username and Password from Client.")
Call wskFile.SendData("UserandPass")
Else:
Call wskFile.SendData("Logedin" & "|" & "Yes")
End If
End Sub

Private Sub wskFile_DataArrival(ByVal bytesTotal As Long) 'what to do when data is recieved...
On Error Resume Next
Dim strData As String
wskFile.GetData strData

Select Case strData
Case "Drive":
    If Permissions = "Yes" Then
        If CheckList(lst4, "Drive") = True Then
        wskFile.SendData ("Perm")
        Exit Sub
        End If
        End If
    If Logedin = "Yes" Then
    Call AddText("Requested drive list.")
    Call DriveList(Drive, wskFile)
    End If

Case "Shutdown":
    If Permissions = "Yes" Then
        If CheckList(lst4, "Shutdown") = True Then
        wskFile.SendData ("Perm")
        Exit Sub
        End If
        End If
    If Logedin = "Yes" Then
    Call Shutdown
    Call AddText("Requested Server shutdown.")
    End If

Case "ProcList":
    If Permissions = "Yes" Then
        If CheckList(lst4, "ProcList") = True Then
        wskFile.SendData ("Perm")
        Exit Sub
        End If
        End If
    If Logedin = "Yes" Then
    Call ListProcess(wskFile)
    Call AddText("Requested process list.")
    End If
    
Case Else: Call DataParsing(strData, wskFile) 'if is isn't a simple, one line command, kick it to
'a module
End Select
End Sub
