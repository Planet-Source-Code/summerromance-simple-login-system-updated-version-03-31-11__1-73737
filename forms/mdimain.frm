VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdimain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Infinity Server Version 1.0"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   11385
   Icon            =   "mdimain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBtn 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   1
      Left            =   2445
      ScaleHeight     =   7425
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   210
      Begin VB.CommandButton btn 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Show Navigation"
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.Timer T_standby 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   6600
   End
   Begin MSComctlLib.StatusBar stabar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7455
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12859
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   635
            MinWidth        =   635
            Picture         =   "mdimain.frx":12025
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "7:17 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   635
            MinWidth        =   635
            Picture         =   "mdimain.frx":12669
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "3/31/2011"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2450
      Begin VB.PictureBox tryicon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   120
         Picture         =   "mdimain.frx":12CAD
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   8
         Top             =   6600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2200
         ScaleHeight     =   345
         ScaleWidth      =   210
         TabIndex        =   3
         Top             =   0
         Width           =   240
         Begin VB.CommandButton btn 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Hide Navigation"
            Top             =   0
            Width           =   210
         End
      End
      Begin MSComctlLib.ImageList iShortcut 
         Left            =   1800
         Top             =   6480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdimain.frx":24CD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdimain.frx":2C616
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdimain.frx":3FE80
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvShortcut 
         Height          =   6735
         Left            =   0
         TabIndex        =   2
         Top             =   400
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483629
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Navigation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   225
         Left            =   720
         TabIndex        =   5
         Top             =   70
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Menu mnustart 
      Caption         =   "&Start"
      Begin VB.Menu mnulogin 
         Caption         =   "&Login"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnulogout 
         Caption         =   "Lo&gout"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnufilebar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Enabled         =   0   'False
      Begin VB.Menu mnusettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnucompact 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mnuchapass 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuacctma 
      Caption         =   "Accout Management"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
      Begin VB.Menu mnuaboutinf 
         Caption         =   "About Infinity"
      End
   End
   Begin VB.Menu mmasry 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
   End
End
Attribute VB_Name = "mdimain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim interval As Integer

Private Sub btn_Click(Index As Integer)
If Index = 0 Then
   Picture1.Visible = False
   picBtn(1).Visible = True
   
ElseIf Index = 1 Then
   Picture1.Visible = True
   picBtn(1).Visible = False

End If

End Sub

Private Sub MDIForm_Load()
Call CheckSoftware(mdimain)
delicon
Call seticon
gHW = Me.hWnd
Hook
App.TaskVisible = False
T_standby.Enabled = True

End Sub

Public Function FOR_LEFT_PICTURE()
Dim lv_listitems As ListItem
Set lvShortcut.SmallIcons = iShortcut
Set lvShortcut.Icons = iShortcut

lvShortcut.ColumnHeaders.Add , , ""
lvShortcut.ColumnHeaders(1).Width = lvShortcut.Width - 80
  With lvShortcut
    .ListItems.Clear
    .ListItems.Add , "L1", "Command Line", 1, 1
    .ListItems.Add , "L2", "Users Information", 2, 2
    .ListItems.Add , "L3", "System Logs", 3, 3
  End With

End Function

Private Sub lvShortcut_DblClick()
    Select Case lvShortcut.SelectedItem.Key
        Case "L1":
                   MinAllExceptOne ("frmCommandLine")
                   frmCommandLine.Show
        Case "L2":
                   MinAllExceptOne ("frmUserInfo")
                   frmUserInfo.Show
        Case "L3":
                   MinAllExceptOne ("frmLogs")
                   frmLogs.Show
        End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Or UnloadMode = 1 Or UnloadMode = 2 Then
   If MsgBox("Are you sure?     ", vbQuestion + vbYesNo) = vbYes Then
      If loginSucceed = True Then
         loginSucceed = False
      End If
      App.TaskVisible = True
      delicon
      Unhook
      Cancel = False
   Else
      Cancel = True
   End If
End If

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
   Me.Hide
   Popup "Infinity Server has been minimize to tray!" & vbCrLf & "right click here to maximize  ", "Infinity Server   "
End If

lvShortcut.Move 0, Shape1.Height + 20, Picture1.Width - 10, Me.ScaleHeight - 410
picBtn(1).Move 0, 0, 210, Me.ScaleHeight
btn(1).Move 0, 0, 210, picBtn(1).Height

End Sub

Private Sub mnuchapass_Click()
frmChangePass.Show 1, Me

End Sub

Private Sub mnucompact_Click()
CompactRepairAccessDB

End Sub

Private Sub mnuexit_Click()
Unload Me

End Sub

Private Sub mnulogin_Click()
frmLogin.Show 1, Me

If loginSucceed = True Then
   mnulogin.Enabled = False
   mnulogout.Enabled = True
   mnuoptions.Enabled = True
   Picture1.Visible = True
   
   FOR_LEFT_PICTURE
   stabar.Panels(1).Text = "Logged in as " & loggedname
End If

End Sub

Private Sub mnulogout_Click()
If MsgBox("Are you sure?   ", vbQuestion + vbYesNo) = vbYes Then
   mnulogin.Enabled = True
   mnulogout.Enabled = False
   mnuoptions.Enabled = False
   Picture1.Visible = False
   picBtn(1).Visible = False
   loginSucceed = False
   lvShortcut.ListItems.Clear
   stabar.Panels(1).Text = "Logged out"
   T_standby.Enabled = True
   Call closeAllExceptOne("")
   Call UserLogs(acctid, "Logged Out")
   
End If

End Sub

Private Sub mnuShow_Click()
Me.WindowState = 2
Me.Show

End Sub

Private Sub T_standby_Timer()
interval = interval + 1

If interval = 15 Then
   stabar.Panels(1).Text = "Stand by "
   T_standby.Enabled = False
   interval = 0
End If

End Sub

Private Sub tryicon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   If Me.WindowState = 1 Then
      Me.PopupMenu Me.mmasry
   End If
   
End If

End Sub
