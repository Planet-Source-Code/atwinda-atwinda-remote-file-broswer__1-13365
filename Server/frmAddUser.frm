VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Add User"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Username + Password :"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4935
      Begin VB.TextBox txtConfrim 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Confrim Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
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
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAddUser.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click() 'makes sure the user input is in the correct format
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

Call SendtoOptions
End Sub

Function SendtoOptions() 'the i did this (is werid) but it changes the listbox, then clears the file, then writes over it again, so this just changed the listbox
Dim User As String, Pass As String, EnPass As String
User = Trim(txtUsername.Text)
Pass = Trim(txtPassword.Text)
EnPass = encryptAll(Pass$, "219119") 'this is a simple encryption in the UserPass Mod
Call frmOptions.lstUsers.AddItem(User)
Call frmOptions.lstPasswords.AddItem(EnPass)
ChangedUsers = "Yes" 'set that somethign was changed
Unload Me
End Function
