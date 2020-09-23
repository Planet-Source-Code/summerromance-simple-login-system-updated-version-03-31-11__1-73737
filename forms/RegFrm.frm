VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form RegFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users Registration"
   ClientHeight    =   5850
   ClientLeft      =   3075
   ClientTop       =   3135
   ClientWidth     =   8535
   Icon            =   "RegFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pct_genInfo 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   8055
      TabIndex        =   18
      Top             =   1440
      Width           =   8055
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   9
         Left            =   5450
         MaxLength       =   11
         TabIndex        =   9
         Top             =   2280
         Width           =   2350
      End
      Begin VB.ComboBox comboInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         IntegralHeight  =   0   'False
         ItemData        =   "RegFrm.frx":1385A
         Left            =   4680
         List            =   "RegFrm.frx":13864
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   10
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   8
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   8
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   5760
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   6
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   1320
         MaxLength       =   18
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   18
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox comboInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         IntegralHeight  =   0   'False
         ItemData        =   "RegFrm.frx":13876
         Left            =   6480
         List            =   "RegFrm.frx":13883
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Index           =   7
         Left            =   7200
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1680
         Width           =   615
      End
      Begin VB.Timer TimerSearch 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7320
         Top             =   120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4200
         TabIndex        =   45
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4080
         TabIndex        =   36
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "(M.I.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   6240
         TabIndex        =   33
         Top             =   2040
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "(FIRST)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4440
         TabIndex        =   32
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "(LAST)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2160
         TabIndex        =   31
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5880
         TabIndex        =   26
         Top             =   2760
         Width           =   390
      End
      Begin VB.Image checksamePic 
         Height          =   345
         Index           =   0
         Left            =   3120
         Top             =   240
         Width           =   345
      End
      Begin VB.Image checksamePic 
         Height          =   345
         Index           =   1
         Left            =   3120
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lbltxt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3465
         TabIndex        =   24
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lbltxt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3465
         TabIndex        =   23
         Top             =   960
         Width           =   45
      End
      Begin VB.Image checksamePic 
         Height          =   345
         Index           =   2
         Left            =   3120
         Top             =   960
         Width           =   345
      End
      Begin VB.Image checksamePic 
         Height          =   345
         Index           =   3
         Left            =   3120
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lbltxt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   1320
         TabIndex        =   22
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "(PREFIX)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   7200
         TabIndex        =   21
         Top             =   2040
         Width           =   585
      End
      Begin VB.Image checksamePic 
         Height          =   345
         Index           =   10
         Left            =   3600
         Top             =   2760
         Width           =   345
      End
      Begin VB.Label lbltxt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3465
         TabIndex        =   20
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label lbltxt 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3465
         TabIndex        =   19
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.PictureBox pct_Restrictions 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   5655
      TabIndex        =   37
      Top             =   1440
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkacctopt 
         Caption         =   "&Change Restriction"
         Height          =   495
         Index           =   5
         Left            =   3120
         TabIndex        =   47
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkacctopt 
         Caption         =   "&Change Level"
         Height          =   495
         Index           =   4
         Left            =   1680
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkacctopt 
         Caption         =   "En&able / Disable Account"
         Height          =   495
         Index           =   2
         Left            =   3120
         TabIndex        =   42
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox chkacctopt 
         Caption         =   "&Edit Account"
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkacctopt 
         Caption         =   "&Delete Account"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkacctopt 
         Caption         =   "&Create Account"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkappopt 
         Caption         =   "Terminate Program"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Account Modification"
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Program Settings"
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   1200
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList imgload 
      Left            =   840
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":138A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":13C43
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":13FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":14377
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":14711
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":14AAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":14E45
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":151DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "[ &Cancel ]"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSav 
      Caption         =   "[ &Save ]"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   960
      Width           =   8535
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   240
      Top             =   5280
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
            Picture         =   "RegFrm.frx":15579
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RegFrm.frx":157A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   48
      Top             =   1080
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General Info   "
      TabPicture(0)   =   "RegFrm.frx":161B9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Restrictions   "
      TabPicture(1)   =   "RegFrm.frx":161D5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maintenance Team And Users Registration Form "
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration Form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   3495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   6600
      Picture         =   "RegFrm.frx":161F1
      Top             =   120
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "RegFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputCounter(10) As Integer, loadingc As Integer, lIndex As Integer, optcnt As Integer
Dim checkOk(3) As Boolean, complete(10) As Boolean, fin(3) As Boolean
Dim idNum As Long
Dim a As String, b As String, c As String

Private Sub cmdCan_Click()
Unload Me

End Sub

Private Sub cmdSav_Click()
If edit = True Then
   Call UserLogs(CDbl(acctid), "Updated ID Number [ " & txtInfo(0).Text & " ]")
   EditUpdate
Else
   Call UserLogs(CDbl(acctid), "Created ID Number [" & txtInfo(0).Text & " ]")
   CreateNew
End If
Unload Me

End Sub

Private Sub EditUpdate()
Dim reguser(4) As ADODB.Recordset
Dim cnt As Integer
If db Is Nothing Then dbconnect
Set reguser(2) = New ADODB.Recordset
With reguser(2)
    .Open ("update T_acct_restric Set CreateU = " & CLng(chkacctopt(0).Value) & ", EditU = " & CLng(chkacctopt(1).Value) & ", DisableU =" & CLng(chkacctopt(2).Value) & ", DelU = " & CLng(chkacctopt(3).Value) & ", ChangeL = " & CLng(chkacctopt(4).Value) & ", ChangeR = " & CLng(chkacctopt(5).Value) & ", TerminateP = " & CLng(chkappopt.Value) & " where IDnumber=" & CDbl(txtInfo(0).Text) & " "), db, 3, 3
Set reguser(2) = Nothing
End With

Set reguser(3) = New ADODB.Recordset
With reguser(3)
    .Open ("update T_users_info Set Fname = '" & txtInfo(4).Text & "', Mname = '" & txtInfo(5).Text & "', Lname = '" & txtInfo(6).Text & "', Prefix = '" & txtInfo(7).Text & "', Gender = '" & comboInfo(0).Text & "', Address = '" & txtInfo(8).Text & "', Email = '" & txtInfo(10).Text & "', ContactNumber = " & CDbl(txtInfo(9).Text) & " where IDnumber=" & txtInfo(0).Text & ""), db, 3, 3
Set reguser(3) = Nothing
End With

Set reguser(4) = New ADODB.Recordset
With reguser(4)
    .Open ("update T_users_acct Set userID = '" & txtInfo(1).Text & "', Ulevel = '" & comboInfo(1).Text & "' where IDnumber=" & CDbl(txtInfo(0).Text) & " "), db, 3, 3
Set reguser(4) = Nothing
End With

End Sub

Private Sub CreateNew()
Dim reguser(4) As ADODB.Recordset
Dim cnt As Integer
dbconnect

EncryptDecrypt (txtInfo(3).Text)
txtInfo(3).Text = temp

Set reguser(1) = New ADODB.Recordset
With reguser(1)
    .Open ("insert into T_ID_numbers values(" & CLng(txtInfo(0).Text) & ", '" & Time & "', '" & Date & "') "), db, 3, 3
Set reguser(1) = Nothing
End With

Set reguser(2) = New ADODB.Recordset
With reguser(2)
    .Open ("insert into T_acct_restric values(" & CLng(txtInfo(0).Text) & "," & CLng(chkacctopt(0).Value) & "," & CLng(chkacctopt(1).Value) & "," & CLng(chkacctopt(2).Value) & "," & CLng(chkacctopt(3).Value) & "," & CLng(chkacctopt(4).Value) & "," & CLng(chkacctopt(5).Value) & "," & CLng(chkappopt.Value) & ") "), db, 3, 3
Set reguser(2) = Nothing
End With

Set reguser(3) = New ADODB.Recordset
With reguser(3)
    .Open ("insert into T_users_info values(" & CLng(txtInfo(0).Text) & ", '" & txtInfo(4).Text & "','" & txtInfo(5).Text & "','" & txtInfo(6).Text & "','" & txtInfo(7).Text & "','" & comboInfo(0).Text & "','" & txtInfo(8).Text & "','" & txtInfo(10).Text & "'," & CDbl(txtInfo(9).Text) & ", '" & Time & "', '" & Date & "') "), db, 3, 3
Set reguser(3) = Nothing
End With

Set reguser(4) = New ADODB.Recordset
With reguser(4)
    .Open ("insert into T_users_acct values(" & CLng(txtInfo(0).Text) & ",'" & txtInfo(1).Text & "','" & txtInfo(3).Text & "','Active','" & comboInfo(1).Text & "', '" & Time & "', '" & Date & "') "), db, 3, 3
Set reguser(4) = Nothing
End With

End Sub

Private Sub comboInfo_Click(Index As Integer)

If comboInfo(1).Text = "Administrator" Then
   For optcnt = 0 To 5
       chkacctopt(optcnt).Value = 1
       chkacctopt(optcnt).Enabled = True
   Next
   chkappopt.Enabled = True
   chkappopt.Value = 1
ElseIf comboInfo(1).Text = "Manager" Then
   For optcnt = 0 To 5
       chkacctopt(optcnt).Value = 0
       chkacctopt(optcnt).Enabled = True
   Next
   chkacctopt(0).Value = 1
   chkacctopt(1).Value = 1
   chkappopt.Enabled = True
   chkappopt.Value = 1
Else
   For optcnt = 0 To 5
       chkacctopt(optcnt).Value = 0
       chkacctopt(optcnt).Enabled = False
   Next
   chkappopt.Value = 0
   chkappopt.Enabled = False
End If
checkcomplete

End Sub

Public Sub loadEdit(ID As Long)
Dim EditInfo As ADODB.Recordset
Set EditInfo = New ADODB.Recordset

With EditInfo
    .Open ("select * from T_users_acct where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    txtInfo(0).Text = .Fields("IDnumber")
    txtInfo(1).Text = .Fields("UserID")
    txtInfo(2).Text = .Fields("UserPassword")
    txtInfo(3).Text = .Fields("UserPassword")
    comboInfo(1).Text = .Fields("ulevel")
    c = .Fields("UserID")
    .Close
    End With
Set EditInfo = Nothing

Set EditInfo = New ADODB.Recordset
With EditInfo
    .Open ("select * from T_users_info where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    txtInfo(6).Text = .Fields("Lname")
    txtInfo(4).Text = .Fields("Fname")
    txtInfo(5).Text = .Fields("Mname")
    txtInfo(7).Text = .Fields("Prefix")
    txtInfo(8).Text = .Fields("Address")
    txtInfo(9).Text = .Fields("ContactNumber")
    txtInfo(10).Text = .Fields("Email")
    comboInfo(0).Text = .Fields("Gender")
    .Close
End With
Set EditInfo = Nothing

Set EditInfo = New ADODB.Recordset
With EditInfo
    .Open ("select * from T_acct_restric where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    
    chkacctopt(0).Value = restrictionopt(CBool(.Fields(1)))
    chkacctopt(1).Value = restrictionopt(CBool(.Fields(2)))
    chkacctopt(2).Value = restrictionopt(CBool(.Fields(3)))
    chkacctopt(3).Value = restrictionopt(CBool(.Fields(4)))
    chkacctopt(4).Value = restrictionopt(CBool(.Fields(5)))
    chkacctopt(5).Value = restrictionopt(CBool(.Fields(6)))
    chkappopt.Value = restrictionopt(CBool(.Fields(7)))
    .Close
End With
Set EditInfo = Nothing

If acctid = CLng(txtInfo(0).Text) Then
   comboInfo(1).Enabled = False
   SSTab1.TabEnabled(1) = False
      
Else
   comboInfo(1).Enabled = checkPermission("combo", 0, CDbl(acctid))
   SSTab1.TabEnabled(1) = checkPermission("tab", 0, CDbl(acctid))
   
End If

checkOk(0) = True
checkOk(1) = True
checkOk(2) = True
checkOk(3) = True
      
checkcomplete

End Sub

Private Sub Form_Load()
If db Is Nothing Then dbconnect

If edit = True Then
  If loggedname = "Administrator" Then
     comboInfo(1).Clear
     comboInfo(1).AddItem "Administrator"
     comboInfo(1).AddItem "Manager"
     comboInfo(1).AddItem "Operator"
     
  ElseIf acctid = frmUserInfo.lblInfo(0).Caption And loggedname = "Manager" Then
     comboInfo(1).Clear
     comboInfo(1).AddItem "Manager"
   
   ElseIf loggedname = "Manager" Then
     comboInfo(1).Clear
     comboInfo(1).AddItem "Manager"
     comboInfo(1).AddItem "Operator"
     
   End If
   
   txtInfo(2).Enabled = False
   txtInfo(3).Enabled = False
 
Else
   If loggedname = "Manager" Then
     comboInfo(1).Clear
     comboInfo(1).AddItem "Manager"
     comboInfo(1).AddItem "Operator"
   
     chkappopt.Enabled = False
     For opcnt = 2 To 5
        chkacctopt(opcnt).Enabled = False
     Next
     SSTab1.TabEnabled(1) = checkPermission("tab", 0, CDbl(acctid))
   
   ElseIf loggedname = "Administrator" Then
     comboInfo(1).Clear
     comboInfo(1).AddItem "Administrator"
     comboInfo(1).AddItem "Manager"
     comboInfo(1).AddItem "Operator"
   
     SSTab1.TabEnabled(1) = checkPermission("tab", 0, CDbl(acctid))
     
   End If
     
   checkcomplete
   idNum = Random(100000, 200000)
   txtInfo(0).Text = idNum
End If

End Sub

Private Function restrictionopt(ResOpt As Boolean) As Integer
If ResOpt = True Then
   restrictionopt = 1
Else
   restrictionopt = 0
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
frmUserInfo.loadinfo
frmUserInfo.loadbutton
Unhook
edit = False

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 1 Then
  pct_genInfo.Visible = True
  pct_Restrictions.Visible = False
Else
  pct_genInfo.Visible = False
  pct_Restrictions.Visible = True

End If

End Sub

Private Sub TimerSearch_Timer()
loadingc = loadingc + 1
lbltxt(lIndex).Caption = "checking if available"
checksamePic(lIndex).Picture = imgload.ListImages(loadingc).ExtractIcon

If loadingc = 8 Then
loadingc = 0
End If

End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 8 Then
   Call valid(Index, "typing")
   Exit Sub
End If

If Index = 0 Then
    Call valid(Index, "typing")
     Select Case UCase$(Chr$(KeyAscii))
        Case 0 To 9
        Case Else
             KeyAscii = 0
    End Select

ElseIf Index = 1 Then
     Call valid(Index, "typing")
     Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case Else
             KeyAscii = 0
    End Select

ElseIf Index = 2 Or Index = 3 Then
     Call valid(Index, "typing")
     Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case 0 To 9
        Case Else
             KeyAscii = 0
    End Select

ElseIf Index >= 4 And Index <= 7 Then
      Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case " "
        Case Else
             KeyAscii = 0
     End Select

ElseIf Index = 8 Then
   Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case 0 To 9
        Case "."
        Case " "
        Case Else
             KeyAscii = 0
       End Select

ElseIf Index = 9 Then
     Select Case UCase$(Chr$(KeyAscii))
        Case 0 To 9
        Case Else
             KeyAscii = 0
    End Select

ElseIf Index = 10 Then
   Call valid(Index, "typing")
   Select Case UCase$(Chr$(KeyAscii))
        Case "A" To "Z"
        Case 0 To 9
        Case "."
        Case "@"
        Case "_"
        Case Else
             KeyAscii = 0
       End Select
End If

End Sub

Private Sub txtInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
checkcomplete
If Index = 1 Then
   If InputCounter(1) < 8 Then
      Call valid(Index, "short")
   Else
      Call valid(Index, "", True)
   End If
End If

If Index = 1 Then
   If txtInfo(Index).Text = "" Then
      Call valid(Index, "emp")
      Exit Sub
   ElseIf InputCounter(Index) < 8 Then
      Call valid(Index, "short")
   Else
      Call valid(Index, "validating")
      checkOk(Index) = checksame(Index, txtInfo(Index).Text)
      If checkOk(Index) = True Then
         Call valid(Index, "", checkOk(Index))
      Else
         Call valid(Index, "", checkOk(Index))
      End If
   End If
End If

If Index = 2 Then
   If InputCounter(2) < 8 Then
      txtInfo(3).Text = ""
      Call valid(3, "typing")
      Call valid(Index, "short")
   ElseIf InputCounter(2) >= 8 Then
      txtInfo(3).Text = ""
      Call valid(Index, "", True)
      Call valid(3, "typing")
   End If
End If

If Index = 3 Then
   If txtInfo(2).Text = "" Then
      Call valid(2, "emp")
      Exit Sub
   ElseIf InputCounter(Index) >= 8 Then
      Call valid(Index, "validating")
      checkOk(2) = checksame(Index, txtInfo(2).Text, txtInfo(Index).Text)
      If checkOk(2) = True Then
         Call valid(Index, "", checkOk(2))
      Else
         Call valid(Index, "", checkOk(2))
      End If
   End If
End If

If Index = 10 Then
   If txtInfo(Index).Text = "" Then
      Call valid(Index, "emp")
      Exit Sub
   Else
      Call valid(Index, "validating")
      checkOk(3) = checksame(Index, txtInfo(Index).Text)
      If checkOk(3) = True Then
         Call valid(Index, "", checkOk(3))
      Else
         Call valid(Index, "", checkOk(3))
      End If
   End If
End If
checkcomplete

End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
      
If Index = 0 Then
   If txtInfo(Index).Text = "" Then
      Call valid(Index, "emp")
      Exit Sub
   Else
      Call valid(Index, "validating")
      checkOk(Index) = checksame(Index, txtInfo(Index).Text)
      If checkOk(Index) = True Then
         Call valid(Index, "", checkOk(Index))
      Else
         Call valid(Index, "", checkOk(Index))
      End If
   End If

ElseIf Index = 1 Then
   If txtInfo(Index).Text = "" Then
      Call valid(Index, "emp")
      Exit Sub
   ElseIf InputCounter(Index) < 8 Then
      Call valid(Index, "short")
   Else
      Call valid(Index, "validating")
      checkOk(Index) = checksame(Index, txtInfo(Index).Text)
      If checkOk(Index) = True Then
         Call valid(Index, "", checkOk(Index))
      Else
         Call valid(Index, "", checkOk(Index))
      End If
   End If
   
ElseIf Index = 2 Then
   If txtInfo(Index).Text = "" Then
      txtInfo(3).Text = ""
      Call valid(Index, "emp")
      Call valid(3, "typing")
      Exit Sub
   ElseIf InputCounter(Index) < 8 Then
      txtInfo(3).Text = ""
      Call valid(3, "typing")
      Call valid(Index, "short")
   ElseIf InputCounter(Index) >= 8 Then
     If txtInfo(3).Text = "" Then
        Exit Sub
     Else
         Call valid(Index, "", True)
         Index = 3
     End If
   End If
  
ElseIf Index = 3 Then
   If txtInfo(2).Text = "" Then
      Call valid(2, "emp")
      Exit Sub
   ElseIf InputCounter(2) < 8 Then
      txtInfo(3).Text = ""
      Call valid(3, "typing")
   ElseIf InputCounter(Index) >= 8 Then
      Call valid(Index, "validating")
      checkOk(2) = checksame(Index, txtInfo(2).Text, txtInfo(Index).Text)
      If checkOk(2) = True Then
         Call valid(Index, "", checkOk(2))
      Else
         Call valid(Index, "", checkOk(2))
      End If
   End If

ElseIf Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Then
   Dim caps As String
   caps = txtInfo(Index).Text
   If txtInfo(Index).Text = "" Then
      Exit Sub
   Else
     Mid$(caps, 1, 1) = UCase$(caps)
     txtInfo(Index).Text = caps
   End If
   
   If txtInfo(7).Text = "" Then
      txtInfo(7).Text = "NA"
   End If

ElseIf Index = 10 Then
   If txtInfo(Index).Text = "" Then
      Call valid(Index, "emp")
      Exit Sub
   Else
      Call valid(Index, "validating")
      checkOk(3) = checksame(Index, txtInfo(Index).Text)
      If checkOk(3) = True Then
         Call valid(Index, "", checkOk(3))
      Else
         Call valid(Index, "", checkOk(3))
      End If
   End If
      
End If
UnSubclassWnd txtInfo(Index).hWnd
checkcomplete

End Sub

Private Sub valid(Index As Integer, ByRef str As String, Optional str1 As Boolean)
   
   If str = "validating" Then
      lIndex = Index
      TimerSearch.Enabled = True

   ElseIf str = "typing" Then
      If Index >= 4 And Index <= 9 Then
         Exit Sub
      Else
         lbltxt(Index).Caption = ""
         checksamePic(Index).Picture = Nothing
      End If
   
   ElseIf str = "emp" Then
       If Index = 1 Then
          lbltxt(Index).Caption = "Username is empty"
       ElseIf Index = 2 Then
          lbltxt(Index).Caption = "Password is empty"
       ElseIf Index = 10 Then
          lbltxt(Index).Caption = "Email Address is empty"
      End If
          checksamePic(Index).Picture = i16x16.ListImages(2).ExtractIcon

   ElseIf str = "short" Then
       lbltxt(Index).Caption = "must contain 8 characters long or 15 characters max"
       checksamePic(Index).Picture = i16x16.ListImages(2).ExtractIcon

   ElseIf str1 = True Then
       lIndex = Index
       TimerSearch.Enabled = False
       lbltxt(Index).Caption = ""
       checksamePic(Index).Picture = i16x16.ListImages(1).ExtractIcon

   ElseIf str1 = False Then
       lIndex = Index
       TimerSearch.Enabled = False
       checksamePic(Index).Picture = i16x16.ListImages(2).ExtractIcon
       If Index = 0 Then
          idNum = Random(100000, 200000)
          txtInfo(0).Text = idNum
       ElseIf Index = 1 Then
          lbltxt(Index).Caption = "Username is not available"
       ElseIf Index = 2 Or Index = 3 Then
          lbltxt(Index).Caption = "Password does not match"
   ElseIf Index = 10 Then
      lbltxt(Index).Caption = "Email Address is not available"
   Else
      lbltxt(Index).Caption = ""
   End If
 End If


End Sub

Private Sub checkcomplete()
Dim compcount  As Integer

For compcount = 0 To 10
    InputCounter(compcount) = Len(Trim(txtInfo(compcount).Text))

If InputCounter(compcount) = 0 Then
   complete(compcount) = False
Else
   complete(compcount) = True
End If

Next

If complete(0) = True And complete(1) = True And complete(2) = True And complete(3) = True And _
   complete(4) = True And complete(5) = True And complete(6) = True And complete(7) = True And _
   complete(8) = True And complete(9) = True And complete(10) = True Then
   fin(0) = True
Else
   fin(0) = False
End If

If checkOk(0) = True And checkOk(1) = True And checkOk(2) = True And checkOk(3) = True Then
   fin(1) = True
Else
   fin(1) = False
End If

If InputCounter(1) < 8 Or InputCounter(2) < 8 Or InputCounter(3) < 8 Then
   fin(2) = False
Else
   fin(2) = True
End If

If comboInfo(0).Text = "" Or comboInfo(1).Text = "" Then
   fin(3) = False
Else
   fin(3) = True
End If

If fin(0) = True And fin(1) = True And fin(2) = True And fin(3) = True Then
   cmdSav.Enabled = True
Else
   cmdSav.Enabled = False
End If

End Sub

Private Function checksame(Index As Integer, st1 As String, Optional st2 As String) As Boolean
Dim chkname As ADODB.Recordset
If db Is Nothing Then dbconnect

If Index = 0 Then
   Set chkname = New ADODB.Recordset
   With chkname
     .Open ("select IDnumber from T_ID_numbers where IDnumber=" & CLng(st1) & ""), db, 3, 3
        If edit = True Then
           checksame = True
        Else
           If .RecordCount = 1 Then
              .MoveFirst
              checksame = False
           Else
              checksame = True
           End If
        End If
     .Close
   Set chkname = Nothing
   End With

ElseIf Index = 1 Then
   Set chkname = New ADODB.Recordset
   With chkname
     .Open ("select userID from T_users_acct where userID='" & st1 & "'"), db, 3, 3
      If edit = True And .RecordCount = 1 Then
         a = .Fields("userID")
         b = c
         checksame = editCheck("userID", "T_users_acct", st1)
      Else
         If .RecordCount = 1 Then
            .MoveFirst
            checksame = False
         Else
            checksame = True
         End If
      End If
     .Close
   Set chkname = Nothing
   End With
   
ElseIf Index = 3 Then
   If st1 = st2 Then
      checksame = True
   Else
      checksame = False
   End If
End If

If Index = 10 Then
   Set chkname = New ADODB.Recordset
   With chkname
   .Open ("select Email from T_users_info where Email='" & st1 & "'"), db, 3, 3
    If edit = True And .RecordCount = 1 Then
       a = .Fields("Email")
       b = frmUserInfo.lvlist(0).SelectedItem.SubItems(7)
       checksame = editCheck("Email", "T_users_info", st1)
    Else
       If .RecordCount = 1 Then
          .MoveFirst
          checksame = False
       Else
          checksame = True
       End If
    End If
     .Close
   Set chkname = Nothing
   End With
End If

End Function

Private Function editCheck(ByRef k As String, ByRef x As String, ByRef z As String) As Boolean
Dim EditRec As ADODB.Recordset
Set EditRec = New ADODB.Recordset

With EditRec
    .Open ("select '" & k & "' from " & x & " where " & k & " = '" & z & "' "), db, 3, 3
    If .RecordCount = 1 Then
        If a = b Then
            editCheck = True
        Else
            editCheck = False
        End If
    Else
       editCheck = True
    End If
    .Close
    Set EditeRec = Nothing
End With

End Function

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   SubClassWnd txtInfo(Index).hWnd
End If

End Sub
