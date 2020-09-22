VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atwinda Remote File: Server Options"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4365
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7699
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Logging"
      TabPicture(0)   =   "frmOptions.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Passwords"
      TabPicture(1)   =   "frmOptions.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Permissions"
      TabPicture(2)   =   "frmOptions.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   " File and Directory Permissions: "
         Height          =   3735
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   5415
         Begin VB.CommandButton cmdunBlocked 
            Caption         =   "<<"
            Height          =   375
            Left            =   2520
            TabIndex        =   25
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmdBlocked 
            Caption         =   ">>"
            Height          =   375
            Left            =   2520
            TabIndex        =   24
            Top             =   960
            Width           =   375
         End
         Begin VB.ListBox lstBlocked 
            Height          =   2460
            IntegralHeight  =   0   'False
            ItemData        =   "frmOptions.frx":0D1E
            Left            =   3000
            List            =   "frmOptions.frx":0D20
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   840
            Width           =   2295
         End
         Begin VB.ListBox lstunBlocked 
            Height          =   2460
            IntegralHeight  =   0   'False
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkPermissions 
            Caption         =   "Yes, I would like file permissions."
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "*You must leave at least one un-blocked command."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3360
            Width           =   3735
         End
         Begin VB.Label Label6 
            Caption         =   "Blocked:"
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Un-Blocked: "
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Password:: "
         Height          =   3735
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   5415
         Begin VB.CommandButton cmdChangePW 
            Caption         =   "Change Password"
            Height          =   255
            Left            =   3720
            TabIndex        =   15
            Top             =   3360
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelUser 
            Caption         =   "Delete User"
            Height          =   255
            Left            =   1920
            TabIndex        =   14
            Top             =   3360
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddUser 
            Caption         =   "Add User"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   3360
            Width           =   1575
         End
         Begin VB.ListBox lstPasswords 
            Height          =   2385
            IntegralHeight  =   0   'False
            Left            =   2760
            TabIndex        =   12
            Top             =   840
            Width           =   2535
         End
         Begin VB.ListBox lstUsers 
            Height          =   2385
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chkPassword 
            Caption         =   "Yes, I want a Username and Password Login"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "Passwords:"
            Height          =   255
            Left            =   2760
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Usernames:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Logging:"
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5415
         Begin VB.CheckBox chkLogging 
            Caption         =   "Yes, I want to enable Logging"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtLogPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label2 
            Caption         =   "File to write log to:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Elements of the program, such as: Logging, Passwords, and file browsing and editing permissions."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLogging_Click() 'if logging is enabled, set the window to look like it
If chkLogging.Value = 1 Then
    txtLogPath.Enabled = True
    If ReadINI("Logging", "Value", App.Path & "\Options.ini") = "Yes" Then
    Logging = "Yes"
    Exit Sub
    Else: Call WriteINI("Logging", "Value", "Yes", App.Path & "\Options.ini")
    End If
Else: txtLogPath.Enabled = False
If ReadINI("Logging", "Value", App.Path & "\Options.ini") = "No" Then
    Logging = "No"
    Exit Sub
    Else: Call WriteINI("Logging", "Value", "No", App.Path & "\Options.ini")
    End If
End If
End Sub

Private Sub chkPassword_Click() 'if password in on, then set that up
If chkPassword = 1 Then
    lstUsers.Enabled = True
    lstUsers.BackColor = &H80000005
    lstPasswords.Enabled = True
    lstPasswords.BackColor = &H80000005
    cmdAddUser.Enabled = True
    cmdDelUser.Enabled = True
    cmdChangePW.Enabled = True
    Call ReadFile(App.Path & "\users.rfd", lstUsers)
    Call ReadFile(App.Path & "\pass.rfd", lstPasswords)
    If ReadINI("UserPass", "Value", App.Path & "\Options.ini") = "Yes" Then
    Exit Sub
    Else: Call WriteINI("UserPass", "Value", "Yes", App.Path & "\Options.ini")
    End If
Else:
    lstUsers.Enabled = False
    lstUsers.BackColor = &H80000000
    lstPasswords.Enabled = False
    lstPasswords.BackColor = &H80000000
    cmdAddUser.Enabled = False
    cmdDelUser.Enabled = False
    cmdChangePW.Enabled = False
    lstUsers.Clear
    lstPasswords.Clear
    If ReadINI("UserPass", "Value", App.Path & "\Options.ini") = "No" Then
    Exit Sub
    Else: Call WriteINI("UserPass", "Value", "No", App.Path & "\Options.ini")
    End If
End If
End Sub

Private Sub chkPermissions_Click() 'if permissions are on, set them up.
If chkPermissions.Value = 1 Then
    lstunBlocked.Enabled = True
    lstunBlocked.BackColor = &H80000005
    lstBlocked.Enabled = True
    lstBlocked.BackColor = &H80000005
    cmdBlocked.Enabled = True
    cmdunBlocked.Enabled = True
    Call ReadFile(App.Path & "\unblocked.rfd", lstunBlocked)
    Call ReadFile(App.Path & "\blocked.rfd", lstBlocked)
    If ReadINI("Permissions", "Value", App.Path & "\Options.ini") = "Yes" Then
    Permissions = "Yes"
    Exit Sub
    Else: Call WriteINI("Permissions", "Value", "Yes", App.Path & "\Options.ini")
    End If
Else:
    lstunBlocked.Enabled = False
    lstunBlocked.BackColor = &H80000000
    lstBlocked.Enabled = False
    lstBlocked.BackColor = &H80000000
    cmdBlocked.Enabled = False
    cmdunBlocked.Enabled = False
    lstunBlocked.Clear
    lstBlocked.Clear
    If ReadINI("Permissions", "Value", App.Path & "\Options.ini") = "No" Then
    Permissions = "No"
    Exit Sub
    Else: Call WriteINI("Permissions", "Value", "No", App.Path & "\Options.ini")
    End If
End If
End Sub

Private Sub cmdAddUser_Click() 'show the adduser form
frmAddUser.Show
End Sub

Private Sub cmdApply_Click() 'if something was changed, then write it to file.
If ChangedUsers = "Yes" Then
Call SaveList(lstUsers, App.Path & "\users.rfd")
DoEvents
Call SaveList(lstPasswords, App.Path & "\pass.rfd")
DoEvents
Call ReadFile(App.Path & "\users.rfd", frmMain.lst1)
Call ReadFile(App.Path & "\pass.rfd", frmMain.lst2)
ChangedUsers = "No"
End If

If ChangedPermissions = "Yes" Then
Call SaveList(lstunBlocked, App.Path & "\unblocked.rfd")
DoEvents
Call SaveList(lstBlocked, App.Path & "\blocked.rfd")
DoEvents
Call ReadFile(App.Path & "\unblocked.rfd", frmMain.lst3)
Call ReadFile(App.Path & "\blocked.rfd", frmMain.lst4)
ChangedPermissions = "No"
End If

If ChangedLogPath = "Yes" Then
If txtLogPath.Text = "" Then
MsgBox "You must enter a path for the log file.", vbOKOnly, "Error"
Exit Sub
Else:
Call WriteINI("Logging", "Path", txtLogPath.Text, App.Path & "\Options.ini")
ChangedLogPath = "No"
End If
End If
End Sub

Private Sub cmdBlocked_Click() 'move the blocked commands to the blocked list
On Error Resume Next
If lstunBlocked.SelCount = 0 Then
Exit Sub
Else:
For i = 0 To lstunBlocked.ListCount - 2
    If lstunBlocked.Selected(i) Then
    Call lstBlocked.AddItem(lstunBlocked.List(i))
    DoEvents
    Call lstunBlocked.RemoveItem(i)
    End If
Next i
End If
ChangedPermissions = "Yes"
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChangePW_Click()
frmChangePW.Show
End Sub

Private Sub cmdDelUser_Click()
frmDeluser.Show
End Sub

Private Sub cmdOK_Click() 'click the apply button, just incase thte user didn't
cmdApply_Click
DoEvents
Unload Me
End Sub

Private Sub cmdunBlocked_Click() 'moves the blocked commands to the unblocked list
On Error Resume Next
If lstBlocked.ListCount = 1 Then
Call lstunBlocked.AddItem(lstBlocked.List(0))
DoEvents
Call lstBlocked.RemoveItem(0)
End If
If lstBlocked.SelCount = 0 Then
Exit Sub
Else:
For i = 0 To lstBlocked.ListCount - 2
    If lstBlocked.Selected(i) Then
    Call lstunBlocked.AddItem(lstBlocked.List(i))
    DoEvents
    Call lstBlocked.RemoveItem(i)
    End If
Next i
End If
ChangedPermissions = "Yes"
End Sub

Private Sub Form_Load() 'set variables and check for settings
ChangedUsers = "No"
ChangePermissions = "No"
ChangedLogPath = "No"
On Error Resume Next
If ReadINI("Logging", "Value", App.Path & "\Options.ini") = "Yes" Then
    If ReadINI("Logging", "Path", App.Path & "\Options.ini") = "" Then
        txtLogPath.Text = "C:\RemoteFileLog.txt"
        ChangedLogPath = "Yes"
    Else
        txtLogPath.Text = LoggingPath
    End If
chkLogging.Value = 1
chkLogging_Click
Else:
chkLogging.Value = 0
chkLogging_Click
End If
If ReadINI("UserPass", "Value", App.Path & "\Options.ini") = "Yes" Then
chkPassword.Value = 1
chkPassword_Click
Else:
chkPassword.Value = 0
chkPassword_Click
End If
If ReadINI("Permissions", "Value", App.Path & "\Options.ini") = "Yes" Then
chkPermissions.Value = 1
chkPermissions_Click
Else:
chkPermissions.Value = 0
chkPermissions_Click
End If
End Sub

Private Sub txtLogPath_Change() 'if the user changes the logging path, tell it to write it...
ChangedLogPath = "Yes"
End Sub
