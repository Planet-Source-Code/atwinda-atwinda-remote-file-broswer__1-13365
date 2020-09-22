VERSION 5.00
Begin VB.Form frmSavebox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Save Box"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5340
   Icon            =   "frmSavebox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "*there is no way to undo this!"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSavebox.frx":0CCA
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSavebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strText

Private Sub cmdSave_Click() 'send the save command to the server
If txtPath.Text > "" And Connected = "Yes" Then
    strText = ""
    strText = "SaveFile" & "|" & txtPath.Text & "š" & txtText.Text
    Call frmMain.wskFile.SendData(strText)
    Unload frmTextlab
    Unload Me
Else: MsgBox "You must enter a path.", vbExclamation, "Error"
End If
End Sub
