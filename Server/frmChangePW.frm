VERSION 5.00
Begin VB.Form frmChangePW 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Change Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Change Password: "
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtConfrim 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Min 5 - Max 20"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Confrim Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Min 5 - Max 15"
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
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "To change a users password. Type in there name below, and then a new password."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click() 'makes sure the correct data was entered...
If txtUsername.Text = "" Or txtPassword.Text = "" Then
    MsgBox "You must enter both Username and Password.", vbOKOnly, "Error"
    Exit Sub
End If
If Len(txtUsername.Text) < 5 Or Len(txtPassword.Text) < 5 Then
    MsgBox "The Username/Password must be more than 5 letters long.", vbOKOnly, "Error"
    Exit Sub
End If
If Not txtPassword.Text = txtConfrim.Text Then
    MsgBox "The password must be the same in both fields.", vbOKOnly, "Error"
    txtPassword.Text = ""
    txtConfrim.Text = ""
    Exit Sub
End If

Call ReplacePW
End Sub

Function ReplacePW() 'finds the username in the username listbox and since it is the same as
'the password list, changes the password in the other box...
Dim User As String, Pass As String, EnPass As String
For i = 0 To frmOptions.lstUsers.ListCount
    If frmOptions.lstUsers.List(i) = Trim(txtUsername.Text) Then
        Call frmOptions.lstPasswords.RemoveItem(i)
        Call frmOptions.lstUsers.RemoveItem(i)
        DoEvents
        User = Trim(txtUsername.Text)
        Pass = Trim(txtPassword.Text)
        EnPass = encryptAll(Pass$, "219119")
        Call frmOptions.lstUsers.AddItem(User)
        Call frmOptions.lstPasswords.AddItem(EnPass)
        End If
Next i
ChangedUsers = "Yes" 'tell it that something was changed...
Unload Me
End Function
