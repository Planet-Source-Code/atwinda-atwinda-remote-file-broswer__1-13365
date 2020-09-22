VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Login"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Username and Password: "
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2295
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
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1215
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
         Left            =   3360
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
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "This Server requires a user name and password to login to it. Please enter your username and password below."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmMain.wskFile.Close
Unload Me
End Sub

Private Sub cmdOK_Click() 'check the length and send them to the server
If txtUser.Text = "" Or txtPass.Text = "" Then
    MsgBox "You must enter both username and password.", vbOKOnly, "Error"
    txtPass.Text = ""
    Exit Sub
End If
If Len(txtUser.Text) < 5 Or Len(txtPass.Text) < 5 Then
    MsgBox "Both your password and username must be more than 5 latters.", vbOKOnly, "Error"
    txtPass.Text = ""
    Exit Sub
End If
Call frmMain.wskFile.SendData("UserPass" & "|" & txtUser.Text & "|" & txtPass.Text)
Call AddText("Sent Username and Password, awaiting reply.")
DoEvents

Unload Me
End Sub
