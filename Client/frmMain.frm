VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Remote File: Client"
   ClientHeight    =   6345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskFile 
      Left            =   1800
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client "
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.OptionButton optcmd 
         Caption         =   "Edit File"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   16
         Top             =   4440
         Width           =   975
      End
      Begin VB.OptionButton optcmd 
         Caption         =   "View File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdProclist 
         Caption         =   "Process List"
         Height          =   255
         Left            =   5880
         TabIndex        =   14
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Shutdown"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         ToolTipText     =   "Shutdown Remote Computer"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdDrive 
         Caption         =   "Drive List"
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         ToolTipText     =   "View Drive Lists"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Display"
         Height          =   255
         Left            =   7320
         TabIndex        =   11
         ToolTipText     =   "Clear Display of Text"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   7440
         TabIndex        =   9
         ToolTipText     =   "Send Command"
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   5280
         Width           =   7215
      End
      Begin VB.OptionButton optcmd 
         Caption         =   "File List"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Use the File List and type in a Files to search in the Command Line."
         Top             =   4680
         Width           =   855
      End
      Begin VB.OptionButton optcmd 
         Caption         =   "Directory List"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Use the Directory List and type in a Directory to search in the Command Line."
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   7440
         TabIndex        =   3
         ToolTipText     =   "Connect to Server"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtDisplay 
         Height          =   3975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   8535
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   8535
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   5520
         Width           =   8535
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5790
         Width           =   615
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   5790
         Width           =   2655
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLabs 
      Caption         =   "Labs"
      Begin VB.Menu mnuTextlab 
         Caption         =   "Text Lab"
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "PopUp Lab"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SvrIP As String
Dim SvrPort As String
Dim SelectCmd As String, Command As String

Private Sub cmdClear_Click()
txtDisplay.Text = ""
End Sub

Private Sub cmdConnect_Click() 'connect to the dns or ip on port 1289...
SvrPort = "1289"
If cmdConnect.Caption = "Connect" Then
    wskFile.Close
    wskFile.Connect SvrIP, SvrPort
    cmdConnect.Caption = "Close"
    lblStatus.Caption = "Connecting..."
ElseIf cmdConnect.Caption = "Close" Then
    wskFile.Close
    cmdConnect.Caption = "Connect"
    lblStatus.Caption = "Closed..."
End If
End Sub

Private Sub cmdDrive_Click() 'send the drive command to the server
If Connected = "Yes" Then
    wskFile.SendData ("Drive")
Else: Call AddText(" Error, No Connection.")
End If
End Sub

Private Sub cmdProclist_Click() 'send the proclist command to the server
If Connected = "Yes" Then
    Call wskFile.SendData("ProcList")
Else: Call AddText(" Error, No Connection.")
End If
End Sub

Private Sub cmdSend_Click() 'check the selected command, and send to the server
If Connected = "Yes" Then
    If txtText.Text > " " Then
        Select Case SelectCmd
        Case "Drive": wskFile.SendData ("Drive")
        Case Else:
            Command = SelectCmd & "|" & txtText.Text
            Call wskFile.SendData(Command)
        End Select
    End If
Else: Call AddText(" Error, No Connection.")
End If
End Sub

Private Sub cmdShutdown_Click() ' send the shutdown command to the server
If Connected = "Yes" Then
    Call wskFile.SendData("Shutdown")
Else: Call AddText(" Error, No Connection.")
End If
End Sub

Private Sub Form_Load() 'enter a dns or IP
SvrIP = InputBox("Please enter a IP or network DNS", "IP/DNS", wskFile.LocalHostName)
End Sub

Private Sub mnuAbout_Click() 'show the about
frmAbout.Show
End Sub

Private Sub mnuPopUp_Click() 'show the popup lab
frmPopUplab.Show
End Sub

Private Sub mnuTextlab_Click() 'show the textlab with the editfile command disabled
frmTextlab.Show
frmTextlab.mnuSaveedited.Enabled = False
End Sub

Private Sub optCmd_Click(Index As Integer) 'select the command based on what is clicked
Select Case Index
Case 0: SelectCmd = "Dir"
Case 1: SelectCmd = "File"
Case 2: SelectCmd = "ViewFile"
Case 3: SelectCmd = "EditFile"
End Select
End Sub

Private Sub wskFile_Close() 'what do if winsock closes
lblStatus.Caption = "Closed..."
cmdConnect.Caption = "Connect"
Connected = "No"
End Sub

Private Sub wskFile_Connect() 'what do when connected
lblStatus.Caption = "Connected..."
Connected = "Yes"
End Sub

Private Sub wskFile_DataArrival(ByVal bytesTotal As Long) 'what do with data when it arrives...
Dim strData As String
wskFile.GetData strData

'these are the simple commands
Select Case strData
Case "UserandPass":
    frmLogin.Show
    cmdSend.Enabled = False
    txtText.Enabled = False
    cmdProclist.Enabled = False
    cmdShutdown.Enabled = False
    cmdDrive.Enabled = False

Case "Perm": Call AddText("This command is not allowed due to server side permissions.")

Case "ServerDown": Call AddText("The remote computer was shutdown. " & Time)

Case "Error": Call AddText("A Server side error occured.")

Case Else: Call DataParsing(strData, wskFile) 'the more 'heavy' commands are sent to the mod for processing
End Select
End Sub
