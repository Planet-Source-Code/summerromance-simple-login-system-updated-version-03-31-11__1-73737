VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserInfo 
   Caption         =   "Users Information"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "frmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   9525
   WindowState     =   2  'Maximized
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   4560
      Visible         =   0   'False
      Width           =   9255
      Begin VB.PictureBox pcbAcct 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   360
         ScaleHeight     =   1455
         ScaleWidth      =   8535
         TabIndex        =   43
         Top             =   1200
         Width           =   8535
         Begin MSComctlLib.ListView lvAcct 
            Height          =   1095
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1931
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID number"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Username"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Account Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Level"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox pcbLogs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   360
         ScaleHeight     =   1455
         ScaleWidth      =   8535
         TabIndex        =   45
         Top             =   1200
         Visible         =   0   'False
         Width           =   8535
         Begin MSComctlLib.ListView lvLogs 
            Height          =   1095
            Left            =   0
            TabIndex        =   46
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1931
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID number"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Activity"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Time Log"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date Log"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSComctlLib.TabStrip tabInfo 
         Height          =   1935
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3413
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Account"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Logs"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSubHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view the list of accounts and their activity logs"
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
         Index           =   1
         Left            =   720
         TabIndex        =   41
         Top             =   480
         Width           =   3945
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of User Accounts And Logs"
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
         Index           =   1
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   2880
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   0
         Picture         =   "frmUserInfo.frx":1385A
         Top             =   120
         Width           =   720
      End
      Begin VB.Shape spMag 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.PictureBox picConn 
      Height          =   135
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   9195
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   9255
   End
   Begin VB.Frame fraList 
      Height          =   4215
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9255
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   4680
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   45
            Width           =   2055
         End
      End
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   2
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox picBtnContainer 
            Height          =   375
            Index           =   6
            Left            =   2160
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   47
            ToolTipText     =   "View Details"
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btn 
               Height          =   315
               Index           =   6
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "View Details"
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2175
            TabIndex        =   3
            Top             =   0
            Width           =   2175
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   5
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   22
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   5
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   4
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   20
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   4
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   21
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   3
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   18
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   19
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   2
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   16
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  ToolTipText     =   "Delete"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   1
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   14
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   15
                  ToolTipText     =   "Edit"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   4
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   5
                  ToolTipText     =   "Create New"
                  Top             =   0
                  Width           =   315
               End
            End
         End
      End
      Begin MSComctlLib.ListView lvlist 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1905
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fistname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Middlename"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Lastname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Prefix"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Gender"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Address"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Email"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Contact Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "CONTACT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   37
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTACT NUMBER."
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
         Index           =   7
         Left            =   5520
         TabIndex        =   36
         Top             =   2040
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   35
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMAIL"
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
         Index           =   6
         Left            =   5520
         TabIndex        =   34
         Top             =   1800
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   33
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADDRESS"
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
         Index           =   5
         Left            =   5520
         TabIndex        =   32
         Top             =   1560
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   31
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GENDER."
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
         Index           =   4
         Left            =   5520
         TabIndex        =   30
         Top             =   1320
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "LAST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LAST NAME"
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
         Index           =   3
         Left            =   960
         TabIndex        =   28
         Top             =   2040
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "MIDDLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIDDLE NAME."
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
         Index           =   2
         Left            =   960
         TabIndex        =   26
         Top             =   1800
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FIRST NAME."
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
         Index           =   1
         Left            =   960
         TabIndex        =   24
         Top             =   1560
         Width           =   3645
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
         TabIndex        =   13
         Top             =   3720
         Width           =   1230
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   0
         Picture         =   "frmUserInfo.frx":2A20C
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Users Information"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lblSubHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view the list of Users and their other Information"
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
         TabIndex        =   11
         Top             =   480
         Width           =   4125
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID NO."
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
         Left            =   960
         TabIndex        =   10
         Top             =   1320
         Width           =   3645
      End
      Begin VB.Label lblname 
         BackColor       =   &H00808080&
         Caption         =   "ID NO."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   765
      End
      Begin VB.Shape spMag 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Accounts and logs"
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":40BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":46E58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8160
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":47AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":484BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":48ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":498E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":4A2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":4AD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserInfo.frx":51568
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim int_size As Integer, int_size_active As Integer
Dim view_other As Boolean
Dim sorts As Boolean
Dim RecCount As Integer

Private Sub btn_Click(Index As Integer)
If db Is Nothing Then dbconnect
Dim checkL As ADODB.Recordset
Set checkL = New ADODB.Recordset

If Index = 0 Then
   Index = 0

ElseIf Index = 3 Then
   Index = 3
   
ElseIf lblInfo(0).Caption = "ID NO." Then
   MsgBox "Please Select A Record!   ", vbExclamation
   Exit Sub

Else
   checkL.Open ("select uLevel from T_users_acct where IDnumber = " & CDbl(lblInfo(0).Caption) & " "), db, 2, 2
End If

Select Case Index
    Case 0
         RegFrm.Show 1, mdimain
    Case 1
         If loggedname = "Manager" And checkL.Fields("uLevel") = "Operator" Then
            edit = True
           Call RegFrm.loadEdit(CLng(frmUserInfo.lvlist(0).SelectedItem.Text))
            RegFrm.Show 1, mdimain
            
         ElseIf loggedname = "Administrator" Then
            edit = True
           Call RegFrm.loadEdit(CLng(frmUserInfo.lvlist(0).SelectedItem.Text))
            RegFrm.Show 1, mdimain
            
         ElseIf acctid = lblInfo(0).Caption Then
            edit = True
           Call RegFrm.loadEdit(CLng(frmUserInfo.lvlist(0).SelectedItem.Text))
            RegFrm.Show 1, mdimain
            
         Else
            MsgBox "You dont have the right to edit this account!   ", vbExclamation
         
            checkL.Close
            Set checkL = Nothing
         
         End If
    Case 2
         If loggedname = "Manager" And checkL.Fields("uLevel") = "Administrator" Then
            MsgBox "You dont have the right to delete this account!   ", vbExclamation
            
            checkL.Close
            Set checkL = Nothing
         
         ElseIf lblInfo(0).Caption = acctid Then
            MsgBox "You can't delete your own account!   ", vbExclamation
            
            checkL.Close
            Set checkL = Nothing
         
         Else
            If MsgBox("Are you sure you want to delete? you wont be able to undo this action!   ", vbQuestion + vbYesNo) = vbYes Then
               Dim del As ADODB.Recordset
               Set del = New ADODB.Recordset
               
               del.Open ("delete from T_ID_numbers WHERE IDnumber = " & CDbl(lvlist(0).SelectedItem.Text) & " "), db, 2, 2
               Call UserLogs(CDbl(acctid), "Deleted ID Number [ " & lvlist(0).SelectedItem.Text & " ]")
               loadinfo
               countrecord
               lvlist(0).SetFocus
               lblSelected(0).Caption = "Selected Record: None"
               Set del = Nothing
               Set db = Nothing
               
            End If
         End If
    Case 3
         loadinfo
    Case 4
    
    Case 5
    
    Case 6
         If view_other = False Then
            AcctInfo = lvlist(0).SelectedItem.Text
         Else
            AcctInfo = lvAcct.SelectedItem.Text
         End If
         frmAcctDetails.Show 1, mdimain
         
End Select

End Sub

Private Sub cmdView_Click()
   If view_other = True And fraList(1).Visible = True Then
        fraList(1).Visible = False
        picConn.Visible = False
        view_other = False
    Else
        fraList(1).Visible = True
        picConn.Visible = True
        view_other = True
        Call lvlist_ItemClick(1, lvlist(0).SelectedItem)
    End If
    int_size = Me.ScaleHeight / 2
    Listview_Resize

End Sub

Private Sub Form_Load()
If db Is Nothing Then dbconnect
SetButtonPicture
loadbutton
loadinfo
countrecord

gHW = Me.hWnd
'Begin subclassing.
Hook

End Sub

Public Sub loadbutton()
btn(0).Enabled = checkPermission("btn", 1, CDbl(acctid))
btn(1).Enabled = checkPermission("btn", 2, CDbl(acctid))
btn(2).Enabled = checkPermission("btn", 4, CDbl(acctid))

End Sub

Public Sub countrecord()
Dim cCount As ADODB.Recordset
Set cCount = New ADODB.Recordset

If db Is Nothing Then dbconnect
cCount.Open ("select IDnumber from T_ID_numbers"), db, 3, 3
RecCount = cCount.RecordCount
cCount.Close
Set cCount = Nothing

If RecCount <= 2 Then
  btn(2).Enabled = False
ElseIf checkPermission("btn", 4, CDbl(acctid)) = False Then
  btn(2).Enabled = False
Else
  btn(2).Enabled = True
End If

lblPageInfo(0).Caption = "0 - 0 of " & RecCount

End Sub

Public Sub loadinfo(Optional sort As String)
Dim userlist As ADODB.Recordset
Set userlist = New ADODB.Recordset

If db Is Nothing Then dbconnect

If sort = "" Then
   sort = "IDnumber Asc"
End If
          
userlist.Open ("select * from T_users_info order by " & sort & " "), db, 3, 3

lvlist(0).ListItems.Clear
lvAcct.ListItems.Clear

Do While Not userlist.EOF
   With userlist
        lvlist(0).ListItems.Add , , !IDnumber, 1, 1
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(1) = !Fname
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(2) = !Mname
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(3) = !Lname
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(4) = !Prefix
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(5) = !Gender
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(6) = !Address
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(7) = !Email
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(8) = !ContactNumber
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(9) = !TimeCreated
        lvlist(0).ListItems(lvlist(0).ListItems.Count).SubItems(10) = !DateCreated
        loadAccts (CDbl(.Fields(0)))
        .MoveNext
    End With
    
Loop
lvlist(0).SelectedItem.Selected = False

userlist.Close
Set userlist = Nothing


End Sub

Private Sub loadAccts(ByRef args As Long)
Dim acctList As ADODB.Recordset
Set acctList = New ADODB.Recordset
   
acctList.Open ("select * from T_users_acct where IDnumber = " & args & " "), db, 3, 3

With acctList
        If .Fields(3) = "Locked" Then
           lvAcct.ListItems.Add , , !IDnumber, 2, 2
     
        Else
           lvAcct.ListItems.Add , , !IDnumber, 1, 1

        End If
        lvAcct.ListItems(lvAcct.ListItems.Count).SubItems(1) = !userid
        lvAcct.ListItems(lvAcct.ListItems.Count).SubItems(2) = !acctstatus
        lvAcct.ListItems(lvAcct.ListItems.Count).SubItems(3) = !ulevel
        lvAcct.ListItems(lvAcct.ListItems.Count).SubItems(4) = !TimeCreated
        lvAcct.ListItems(lvAcct.ListItems.Count).SubItems(5) = !DateCreated
'        .MoveNext
    .Close
End With
H_Locked_Acct
Set acctList = Nothing

End Sub

Private Sub H_Locked_Acct()
Dim Items As MSComctlLib.ListItem
Dim Locked_ID As String
Dim cols As Integer

For Each Items In lvAcct.ListItems
    If Items.SubItems(2) = "Locked" Then
       For cols = 1 To 5
           Items.ForeColor = &HFF&
           Items.ListSubItems(cols).ForeColor = &HFF&
       Next
    Locked_ID = Items.Text
    End If
Next

For Each Items In lvlist(0).ListItems
    If Items.Text = Locked_ID Then
       For cols = 1 To 10
           Items.ForeColor = &HFF&
           Items.ListSubItems(cols).ForeColor = &HFF&
       Next
    End If
Next

End Sub

Private Sub SetButtonPicture()
    With Me
        btn(0).Picture = .i16x16.ListImages(2).ExtractIcon
        btn(1).Picture = .i16x16.ListImages(3).ExtractIcon
        btn(2).Picture = .i16x16.ListImages(4).ExtractIcon
        btn(3).Picture = .i16x16.ListImages(5).ExtractIcon
        btn(4).Picture = .i16x16.ListImages(1).ExtractIcon
        btn(5).Picture = .i16x16.ListImages(6).ExtractIcon
        btn(6).Picture = .i16x16.ListImages(7).ExtractIcon
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
view_other = False
Set db = Nothing
Unhook

End Sub

Private Sub Form_Resize()
int_size = Me.ScaleHeight / 2
Listview_Resize
 
End Sub

Private Sub picConn_DblClick()
    int_size = Me.ScaleHeight / 2
    Listview_Resize

End Sub

Private Sub picConn_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    int_size_active = 1
End Sub

Private Sub picConn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If int_size_active = 1 Then
        If y < 0 Then
            If picConn.Top > (fraList(0).Top + 2600) Then
                picConn.Top = picConn.Top - (-(y))
            End If
        Else
            If fraList(1).Height >= 2600 Then
                picConn.Top = picConn.Top + y
            End If
        End If
        int_size = picConn.Top
        Listview_Resize
    End If

End Sub

Private Sub picConn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    int_size_active = 0
End Sub

Private Sub Listview_Resize()
On Error Resume Next
    Dim i As Integer
    If view_other = False Then
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, Me.ScaleHeight - (cmdView.Height + 180)
        lvlist(i).Move 120, 2320, fraList(i).Width - 240, fraList(i).Height - (2300 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
        lvlist(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        cmdView.Move Me.Width - 2800, Me.ScaleHeight - 500
     Else
        picConn.Width = Me.ScaleWidth - 240
        picConn.Top = int_size
        'this will use for listview(0)
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, picConn.Top - (105)
        lvlist(i).Move 120, 2320, fraList(i).Width - 240, fraList(i).Height - (2300 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 120)
        lvlist(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        'this will use for listview(1)
        i = 1
        fraList(i).Move 120, (picConn.Top + picConn.Height) - 15, Me.ScaleWidth - 240, (Me.ScaleHeight - ((picConn.Top - 15) + 240 + 240 + 120 + 80))
        spMag(i).Move 0, 240, fraList(i).Width
        tabInfo.Move 120, 840, fraList(i).Width - 214, fraList(i).Height - 1100
        pcbAcct.Move tabInfo.Left + 50, tabInfo.Top + 350, tabInfo.Width - 80, tabInfo.Height - 400
        pcbLogs.Move tabInfo.Left + 50, tabInfo.Top + 350, tabInfo.Width - 80, tabInfo.Height - 400
        lvAcct.Move 120, 200, pcbAcct.Width - 300, pcbAcct.Height - 300
        lvLogs.Move 120, 200, pcbAcct.Width - 300, pcbAcct.Height - 300
        cmdView.Move Me.Width - 2800, Me.ScaleHeight - 500
    End If
End Sub

Private Sub lvlist_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If Index = 0 Then
   If ColumnHeader.Index = 1 Then
      If sorts = False Then
         sorts = True
         loadinfo ("IDnumber Desc")
      Else
         sorts = False
         loadinfo ("IDnumber Asc")
      End If
   ElseIf ColumnHeader.Index = 2 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Fname Desc")
      Else
         sorts = False
         loadinfo ("Fname Asc")
      End If
   ElseIf ColumnHeader.Index = 3 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Mname Desc")
      Else
         sorts = False
         loadinfo ("Mname Asc")
      End If
   ElseIf ColumnHeader.Index = 4 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Lname Desc")
      Else
         sorts = False
         loadinfo ("Lname Asc")
      End If
   ElseIf ColumnHeader.Index = 5 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Prefix Desc")
      Else
         sorts = False
         loadinfo ("Prefix Asc")
      End If
   ElseIf ColumnHeader.Index = 6 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Gender Desc")
      Else
         sorts = False
         loadinfo ("Gender Asc")
      End If
   ElseIf ColumnHeader.Index = 7 Then
      If sorts = False Then
         sorts = True
         loadinfo ("Address Desc")
      Else
         sorts = False
         loadinfo ("Address Asc")
     End If
   ElseIf ColumnHeader.Index = 8 Then
     If sorts = False Then
        sorts = True
        loadinfo ("Email Desc")
     Else
        sorts = False
        loadinfo ("Email Asc")
     End If
   ElseIf ColumnHeader.Index = 9 Then
     If sorts = False Then
        sorts = True
        loadinfo ("ContactNumber Desc")
     Else
        sorts = False
        loadinfo ("ContactNumber Asc")
     End If
   ElseIf ColumnHeader.Index = 10 Then
     If sorts = False Then
        sorts = True
        loadinfo ("TimeCreated Desc")
     Else
        sorts = False
        loadinfo ("TimeCreated Asc")
     End If
   ElseIf ColumnHeader.Index = 11 Then
     If sorts = False Then
        sorts = True
        loadinfo ("DateCreated Desc")
     Else
        sorts = False
        loadinfo ("DateCreated Asc")
     End If
   Else
     Exit Sub
   End If
End If

End Sub

Private Sub lvlist_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim Value As String
Dim listcount As Integer
listcount = lvlist(0).ListItems.Count

If listcount = 0 Then
   Exit Sub
Else
   lblInfo(0).Caption = Item.Text
   lblInfo(1).Caption = Item.SubItems(1)
   lblInfo(2).Caption = Item.SubItems(2)
   lblInfo(3).Caption = Item.SubItems(3)
   lblInfo(4).Caption = Item.SubItems(5)
   lblInfo(5).Caption = Item.SubItems(6)
   lblInfo(6).Caption = Item.SubItems(7)
   lblInfo(7).Caption = Item.SubItems(8)
   
   lblSelected(0).Caption = "Selected Record: " & Item.Index
   lblPageInfo(0).Caption = Item.Index & " - " & listcount & " of " & RecCount

   If view_other = False Then
      Exit Sub
   Else
      Value = Item.Text
      For Each Item In lvAcct.ListItems
        If LCase(Item) = LCase(Value) Then 'find complete words
          Item.Selected = True
          Item.EnsureVisible
          If pcbAcct.Visible = True Then
             lvAcct.SetFocus
          Else
             lvLogs.SetFocus
          End If
          Exit For
        End If
        Next
        Call lvAcct_ItemClick(lvlist(0).SelectedItem)
   End If
End If

End Sub

Private Sub lvAcct_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim ulog As ADODB.Recordset
Set ulog = New ADODB.Recordset
If db Is Nothing Then dbconnect
ulog.Open ("select * from T_user_logs where IDnumber = " & CDbl(Item.Text) & " "), db, 3, 3

With ulog
   If .RecordCount > 0 Then
     lvLogs.ListItems.Clear
     Do While Not .EOF Or .BOF
        lvLogs.ListItems.Add , , !IDnumber, 1, 1
        lvLogs.ListItems(lvLogs.ListItems.Count).SubItems(1) = !Activities
        lvLogs.ListItems(lvLogs.ListItems.Count).SubItems(2) = !TimeLog
        lvLogs.ListItems(lvLogs.ListItems.Count).SubItems(3) = !DateLog
        .MoveNext
     Loop
   Else
      lvLogs.ListItems.Clear
   End If
   .Close
   Set ulog = Nothing
End With

End Sub

Private Sub tabInfo_Click()
If tabInfo.SelectedItem.Index = 1 Then
   pcbAcct.Visible = True
   pcbLogs.Visible = False
   lvAcct.SetFocus
Else
   pcbAcct.Visible = False
   pcbLogs.Visible = True
   lvLogs.SetFocus
End If

End Sub

