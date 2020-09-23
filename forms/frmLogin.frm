VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4005
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLog 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1440
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1465
         Width           =   2535
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
            Left            =   120
            TabIndex        =   4
            Top             =   120
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
            Left            =   1320
            TabIndex        =   5
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   30
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
            Index           =   0
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
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
            Index           =   1
            Left            =   1560
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboAcc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmLogin.frx":0000
            Left            =   1560
            List            =   "frmLogin.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Username:"
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
            Left            =   720
            TabIndex        =   9
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pa&ssword:"
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
            Left            =   720
            TabIndex        =   8
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
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
            Left            =   720
            TabIndex        =   7
            Top             =   960
            Width           =   555
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   0
            Picture         =   "frmLogin.frx":0004
            Top             =   195
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOk_Click()
If txtinput(0).Text = "" Then
   MsgBox "Please type in your Username    ", vbInformation
   txtinput(0).SetFocus

ElseIf txtinput(1).Text = "" Then
   MsgBox "Please type in your Password    ", vbInformation
   txtinput(1).SetFocus
   
Else
   Call authenticate(txtinput(0).Text, txtinput(1).Text)
   If loginSucceed = True Then
      Unload Me
   Else
     If validuser(0) = False Then
        txtinput(0).Text = ""
        txtinput(1).Text = ""
        txtinput(0).SetFocus
        Exit Sub
     ElseIf userlocked = True Then
        txtinput(0).Text = ""
        txtinput(1).Text = ""
        txtinput(0).SetFocus
        Exit Sub
     Else
        loginfail = loginfail + 1
        Call loginfailcheck(txtinput(0).Text)
        txtinput(1).Text = ""
     End If
   End If
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnSubclassWnd txtinput(0).hwnd
UnSubclassWnd txtinput(1).hwnd

End Sub

Private Sub txtinput_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 8 Then Exit Sub
End If

If Index = 0 Then
     Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case Else
             KeyAscii = 0
     End Select
   
ElseIf Index = 1 Then
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
   SubClassWnd txtinput(Index).hwnd
End If

End Sub
