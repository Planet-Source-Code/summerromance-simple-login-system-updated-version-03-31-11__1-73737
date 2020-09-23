VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogs 
   Caption         =   "List of System Logs"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "frmLogs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   9555
   WindowState     =   2  'Maximized
   Begin VB.Frame fraList 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   3
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2175
            TabIndex        =   4
            Top             =   0
            Width           =   2175
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   2
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   11
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   "Delete"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   3
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   9
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   4
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   7
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   4
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   5
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   5
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   5
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   6
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   4680
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   1
         Top             =   3720
         Width           =   4545
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   2
            Top             =   45
            Width           =   2055
         End
      End
      Begin MSComctlLib.ListView lvlist 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1905
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Activities"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time Logs"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Logs"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblSubHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view the list of Program logs"
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
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   480
         Width           =   2640
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Logs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   1140
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   0
         Picture         =   "frmLogs.frx":17703
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Record:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1230
      End
      Begin VB.Shape spMag 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   9255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":2E0B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8040
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":3434F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":34D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":35773
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":36185
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":36B97
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogs.frx":375A9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_size As Integer
Dim view_other As Boolean
Dim sorts As Boolean
Dim RecCount As Integer

Private Sub btn_Click(Index As Integer)
If db Is Nothing Then dbconnect
Dim checkL As ADODB.Recordset
Set checkL = New ADODB.Recordset

checkL.Open ("select uLevel from T_users_acct where IDnumber = " & CDbl(lvlist(0).SelectedItem.Text) & " "), db, 2, 2

Select Case Index
    Case 2
         If loggedname = "Administrator" Then
            If MsgBox("Are you sure you want to delete? you wont be able to undo this action!   ", vbQuestion + vbYesNo) = vbYes Then
               Dim del As ADODB.Recordset
               Set del = New ADODB.Recordset
               
               del.Open ("delete from T_user_logs WHERE IDnumber = " & CDbl(lvlist(0).SelectedItem.Text) & " "), db, 2, 2
               Call UserLogs(CDbl(acctid), "Deleted ID Number [ " & lvlist(0).SelectedItem.Text & " ]")
               
               loadinfo
               countrecord
               lvlist(0).SetFocus
               lblSelected(0).Caption = "Selected Record: None"
               Set del = Nothing
               
            End If
            
         Else
            MsgBox "You dont have the right to delete this log!   ", vbExclamation
            
            checkL.Close
            Set checkL = Nothing
         
         End If
    Case 3
         loadinfo
    Case 4
    
    Case 5
    
End Select

End Sub

Private Sub Form_Load()
SetButtonPicture
loadbutton
loadinfo
countrecord
gHW = Me.hwnd
'Begin subclassing.
Hook

End Sub

Public Sub loadbutton()
btn(2).Enabled = checkPermission("btn", 3, CDbl(acctid))

End Sub

Public Sub countrecord()
Dim cCount As ADODB.Recordset
Set cCount = New ADODB.Recordset

If db Is Nothing Then dbconnect
cCount.Open ("select IDnumber from T_user_logs"), db, 3, 3
RecCount = cCount.RecordCount
cCount.Close
Set cCount = Nothing

If RecCount <= 2 Or checkPermission("btn", 3, CDbl(acctid)) = False Then
  btn(2).Enabled = False
Else
  btn(2).Enabled = True
End If

lblPageInfo(0).Caption = "0 - 0 of " & RecCount

End Sub

Public Sub loadinfo(Optional sort As String)
Dim UserLogs As ADODB.Recordset
Set UserLogs = New ADODB.Recordset

If db Is Nothing Then dbconnect

If sort = "" Then
   sort = "IDnumber Asc"
End If
          
UserLogs.Open ("select * from T_user_logs order by " & sort & " "), db, 3, 3

lvlist(0).ListItems.Clear

Do While Not UserLogs.EOF
   With UserLogs
        lvlist(0).ListItems.Add , , !IDnumber, 1, 1
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(1) = !Activities
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(2) = !TimeLog
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(3) = !DateLog
        .MoveNext
    End With
    
Loop

UserLogs.Close
Set UserLogs = Nothing


End Sub

Private Sub SetButtonPicture()
    With Me
        btn(2).Picture = .i16x16.ListImages(4).ExtractIcon
        btn(3).Picture = .i16x16.ListImages(5).ExtractIcon
        btn(4).Picture = .i16x16.ListImages(1).ExtractIcon
        btn(5).Picture = .i16x16.ListImages(6).ExtractIcon

    End With
    
End Sub

Private Sub Form_Resize()
int_size = Me.ScaleHeight / 2
Listview_Resize
 
End Sub

Private Sub Listview_Resize()
On Error Resume Next
    Dim i As Integer
    If view_other = False Then
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, Me.ScaleHeight - 300
        lvlist(i).Move 120, picData(0).Height + 1100, fraList(i).Width - 240, fraList(i).Height - (1400 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
        lvlist(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        cmdView.Move 120, Me.ScaleHeight - 400
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unhook

End Sub

Private Sub lvlist_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim Value As String
Dim listcount As Integer
listcount = lvlist(0).ListItems.Count

If listcount = 0 Then
   Exit Sub
Else
   lblSelected(0).Caption = "Selected Record: " & Item.Index
   lblPageInfo(0).Caption = Item.Index & " - " & listcount & " of " & RecCount
   
End If

End Sub
