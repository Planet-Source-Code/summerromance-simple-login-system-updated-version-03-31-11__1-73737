VERSION 5.00
Begin VB.Form frmLock_Unlocl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "frmUnlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtinput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmLock_Unlocl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOk_Click()
If txtinput(0).Text = "" Then
   MsgBox "Please type your password!   ", vbInformation
   txtinput(0).Text = ""
   txtinput(0).SetFocus
      
Else
   Call Lock_UnlockAcct(acctid, txtinput(0).Text, frmAcctDetails.lblInfo(0).Caption, actL)
   
   If actL = "Active" Then
      Call UserLogs(CDbl(acctid), "Activated ID Number [ " & frmAcctDetails.lblInfo(0).Caption & " ]")
   Else
      Call UserLogs(CDbl(acctid), "Locked ID Number [ " & frmAcctDetails.lblInfo(0).Caption & " ]")
   End If
   
   frmAcctDetails.listinfo (frmUserInfo.lblInfo(0).Caption)
   frmAcctDetails.loadbutton
   frmUserInfo.loadinfo
   Unload Me
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
UnSubclassWnd txtinput(0).hWnd

End Sub

Private Sub txtinput_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 8 Then Exit Sub
End If

If Index = 0 Then
  Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case 0 To 9
        Case Else
             KeyAscii = 0
     End Select
End If

End Sub

Private Sub txtinput_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   SubClassWnd txtinput(Index).hWnd
End If

End Sub
