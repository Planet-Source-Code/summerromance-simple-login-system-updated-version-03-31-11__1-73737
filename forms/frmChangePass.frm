VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3945
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
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
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   80
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
         Index           =   2
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   2295
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
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   2295
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
         Index           =   0
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re - Type"
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
         TabIndex        =   8
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Pa&ssword:"
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
         TabIndex        =   7
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOk_Click()
Dim inputCount As Integer
inputCount = Len(Trim(txtinput(2).Text))

If txtinput(0).Text = "" Then
   MsgBox "Please type your old password!   ", vbInformation
   txtinput(0).SetFocus

ElseIf txtinput(1) <> txtinput(2) Then
   MsgBox "New Password does not match!   ", vbCritical
   txtinput(2).SetFocus
   
ElseIf inputCount < 8 Then
   MsgBox "Password must contain 8 characters long or 15 characters max!   ", vbInformation
   txtinput(1).SetFocus
      
Else
   Call changePass(txtinput(0).Text, txtinput(2).Text)
End If

End Sub

Private Sub changePass(oldP As String, newP As String)
Dim userChangeP As ADODB.Recordset
Set userChangeP = New ADODB.Recordset

With userChangeP
     .Open ("select IDnumber, userID, userPassword from T_users_acct where IDnumber = " & CDbl(acctid) & " "), db, 3, 3
     EncryptDecrypt (oldP)
     oldP = temp
     If oldP = .Fields("userPassword") Then
        Set userChangeP = Nothing
        .Close
        
        Set userChangeP = New ADODB.Recordset
        EncryptDecrypt (newP)
        newP = temp
        .Open ("update T_users_acct set userPassword = '" & newP & "' where IDnumber = " & CDbl(acctid) & " "), db, 3, 3
        Set userChangeP = Nothing
        
        MsgBox "Password Change Successful!   ", vbInformation
        Unload Me
        
     Else
        MsgBox "Invalid Password!   ", vbCritical
        txtinput(0).Text = ""
        txtinput(1).Text = ""
        txtinput(2).Text = ""
        txtinput(0).SetFocus
        
     End If
    
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
UnSubclassWnd txtinput(0).hwnd
UnSubclassWnd txtinput(1).hwnd
UnSubclassWnd txtinput(2).hwnd

End Sub

Private Sub txtinput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 Then
    txtinput(2).Text = ""
End If

End Sub

Private Sub txtinput_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 8 Then Exit Sub
End If

If Index < 2 Then
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

