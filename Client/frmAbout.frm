VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: About"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   1155
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote File Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":0000
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   120
      Picture         =   "frmAbout.frx":0CCA
      Top             =   45
      Width           =   6150
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote File Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   990
      TabIndex        =   1
      Top             =   870
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote File Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1020
      TabIndex        =   2
      Top             =   900
      Width           =   4815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblAuthor.Caption = "Andy Stagg of Atwinda Software"
lblWebsite.Caption = "http://atwindasoft.myqth.com"
lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub lblWebsite_Click()
Shell "Explorer " & "http://atwindasoft.myqth.com", vbMaximizedFocus
End Sub
