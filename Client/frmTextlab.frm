VERSION 5.00
Begin VB.Form frmTextlab 
   Caption         =   "Atwinda Remote File: Text Lab"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7935
   Icon            =   "frmTextlab.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtText 
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Begin VB.Menu mnuSavenew 
            Caption         =   "Save As..."
         End
         Begin VB.Menu mnuSaveedited 
            Caption         =   "Save Edited..."
         End
      End
      Begin VB.Menu spacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Text Lab"
      End
   End
End
Attribute VB_Name = "frmTextlab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If Me.Width < 5000 Or Me.Height < 5000 Then
    Me.Width = 5000
    Me.Height = 5000
End If
txtText.Width = Me.Width - (3 * txtText.Left)
txtText.Height = Me.Height - (8 * txtText.Top) + 60
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuSaveedited_Click() 'save an edited file
frmSavebox.txtPath.Text = txtPath.Text
frmSavebox.txtPath.Locked = True
frmSavebox.txtText.Text = txtText.Text
frmSavebox.Show
End Sub

Private Sub mnuSavenew_Click() 'save a new file
frmSavebox.txtText = txtText.Text
frmSavebox.Show
End Sub
