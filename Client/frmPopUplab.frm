VERSION 5.00
Begin VB.Form frmPopUplab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Atwinda Remote File: Popup Lab"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   " Style: "
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   4095
      Begin VB.OptionButton optType 
         Caption         =   "System Modal"
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Question"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optType 
         Caption         =   "Exclamation"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H8000000A&
         Caption         =   "Critical"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H8000000A&
         Caption         =   "Information"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H8000000A&
         Caption         =   "OK Only"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   " Title: "
      Height          =   640
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   " Message:"
      Height          =   980
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
      Begin VB.CommandButton cmdMultiline 
         Caption         =   "Multi line"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtMultiline 
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H8000000B&
      Caption         =   "Send PopUp"
      Height          =   255
      Left            =   2745
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000B&
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "The PopUp Lab allows you to customize the popup box. Style, Title, and Messages are the three catagories."
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmPopUplab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MessageType As String

Private Sub cmdMultiline_Click() 'changes from single to multiline...
If cmdMultiline.Caption = "Multi line" Then
    Me.Height = "4785"
    Frame3.Height = "1680"
    txtMultiline.Height = "1005"
    txtMultiline.Visible = True
    txtMessage.Visible = False
    txtMultiline.Text = ""
    cmdMultiline.Top = "1320"
    cmdCancel.Top = "4080"
    cmdSend.Top = "4080"
    cmdMultiline.Caption = "Single line"
ElseIf cmdMultiline.Caption = "Single line" Then
    Me.Height = "4020"
    Frame3.Height = "980"
    txtMultiline.Height = "285"
    txtMultiline.Visible = False
    txtMessage.Visible = True
    txtMessage.Text = ""
    cmdMultiline.Top = "600"
    cmdCancel.Top = "3360"
    cmdSend.Top = "3360"
    cmdMultiline.Caption = "Multi line"
End If
End Sub

Private Sub cmdSend_Click() ' put together the three different elements, and send 'em
If Connected = "Yes" Then
    If txtTitle > "" Then
        Select Case cmdMultiline.Caption
        Case "Multi line":
        Call frmMain.wskFile.SendData("PopUp" & "|" & txtMessage.Text & "|" & MessageType & "|" & txtTitle.Text)
        Case "Single line":
        Call frmMain.wskFile.SendData("PopUp" & "|" & txtMultiline.Text & "|" & MessageType & "|" & txtTitle.Text)
        End Select
        Call AddText("Sent PopUp Message.")
        Unload Me
    End If
End If
End Sub

Private Sub optType_Click(Index As Integer) 'select the type on box
Select Case Index
Case 0: MessageType = "vbCritical"
Case 1: MessageType = "vbExclamation"
Case 2: MessageType = "vbInformation"
Case 3: MessageType = "vbQuestion"
Case 4: MessageType = "vbOKOnly"
Case 5: MessageType = "vbSystemModal"
End Select
End Sub
