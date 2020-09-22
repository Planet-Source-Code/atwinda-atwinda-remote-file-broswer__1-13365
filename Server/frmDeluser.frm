VERSION 5.00
Begin VB.Form frmDeluser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Delete User"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Delete Username: "
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   2535
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
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Username to Delete:"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "To Delete a user, type there name in below and press OK."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmDeluser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click() ' find the users name, and delete it from the listbox,
'delete the same index from the password list...
If txtUsername.Text = "" Or Len(txtUsername.Text) < 5 Then
MsgBox "You must enter a Username.", vbOKOnly, "Error"
End If

For i = 0 To frmOptions.lstUsers.ListCount
    If frmOptions.lstUsers.List(i) = Trim(txtUsername.Text) Then
        Call frmOptions.lstPasswords.RemoveItem(i)
        Call frmOptions.lstUsers.RemoveItem(i)
    End If
Next i
ChangedUsers = "Yes"
Unload Me
End Sub

