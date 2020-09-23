VERSION 5.00
Begin VB.Form frmCommandLine 
   AutoRedraw      =   -1  'True
   Caption         =   "Infinity Command Line"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   DrawStyle       =   6  'Inside Solid
   Icon            =   "cmd.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox output 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
   Begin VB.TextBox inputcomm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      MaxLength       =   40
      TabIndex        =   0
      Top             =   4920
      Width           =   9375
   End
End
Attribute VB_Name = "frmCommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
output.Text = "-->  Infinity Command Line [ Version 1.0.0 ]" & vbCrLf & "(C) Copyright 2011 - 2015 y-five development" & vbCrLf & vbCrLf
End Sub

Private Sub Form_Resize()
On Error Resume Next
output.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 500
inputcomm.Move 0, output.Height, Me.ScaleWidth - 350
End Sub

Private Sub inputcomm_KeyPress(KeyAscii As Integer)
Dim comms() As String, params() As String
Dim com As String
com = Replace$(LCase$(inputcomm.Text), " ", "")

If KeyAscii = 13 Then
   If com = "cls" Then
          output.Text = ""
   ElseIf com = "exit" Then
          End
   ElseIf com = "users" Or com = "usershelp" Or com = "users/?" Or com = "users/id" Then
          output.Text = output.Text & _
                 "-> Command parameters users [ /list /id /status /activate /lock ]" & vbCrLf & _
                 "   /list            Displays the list of all the users" & vbCrLf & _
                 "   /status       Displays the status of the user account" & vbCrLf & _
                 "   /activate    Activates a blocked account" & vbCrLf & _
                 "   /lock           Locks a user account" & vbCrLf & vbCrLf & _
                 "   Ex users /id [ID NUMBER] /parameter [ status, activate or lock ]" & vbCrLf & _
                 "   Ex users /list" & vbCrLf & vbCrLf
                 
   ElseIf com = "users/list" Then
          comms = Split(com, "/")
          MsgBox comms(0) & " " & comms(1)

   ElseIf com Like "users/id*/status" Or com Like "users/id*/activate" Or com Like "users/id*/lock" Then
            comms = Split(com, "/")
            params = Split(comms(1), ":")
            Mid$(comms(1), 2, 1) = ":"
            params = Split(comms(1), ":")
            Call userscomm(comms(2), params(1))
   Else
        output.Text = output.Text & "->'" & com & "'" & " is not recognized as internal or external command " & vbCrLf & vbCrLf
   End If
inputcomm.Text = ""
End If

End Sub

Public Sub userscomm(ByVal synt As String, ByVal idNum As String)
If idNum = "" Then
   MsgBox "fuck"

Else
Dim user(2) As ADODB.Recordset
dbconnect
    Set user(1) = New ADODB.Recordset
        user(1).Open ("select * from T_users_acct where IDnumber ='" & idNum & "'"), db, 3, 3
           If synt = "status" Then
              output.Text = output.Text & "->'" & "User Account " & user(1).Fields(2) & " with ID number " & idNum & " is " & user(1).Fields(4) & vbCrLf & vbCrLf
           
           ElseIf synt = "lock" Then
              Set user(2) = New ADODB.Recordset
              user(2).Open ("Update T_users_acct set acctStatus = 'Locked' where userID='" & user(1).Fields(2) & "'"), db, 3, 3
              output.Text = output.Text & "->'" & "User Account " & user(1).Fields(2) & " with ID number " & idNum & " Locked" & vbCrLf & vbCrLf
    
           ElseIf synt = "activate" Then
              Set user(2) = New ADODB.Recordset
              user(2).Open ("Update T_users_acct set acctStatus = 'Active' where userID='" & user(1).Fields(2) & "'"), db, 3, 3
              output.Text = output.Text & "->'" & "User Account " & user(1).Fields(2) & " with ID number " & idNum & " Activated" & vbCrLf & vbCrLf
              
           End If
user(1).Close
Set user(1) = Nothing
Set user(2) = Nothing

End If

End Sub
