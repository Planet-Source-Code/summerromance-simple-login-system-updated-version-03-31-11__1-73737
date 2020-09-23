VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAcctDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Details"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9765
   Icon            =   "frmAcctDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList i16x16 
      Left            =   240
      Top             =   5160
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
            Picture         =   "frmAcctDetails.frx":1385A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctDetails.frx":1426C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctDetails.frx":1753D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctDetails.frx":19D5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctDetails.frx":1A771
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAcctDetails.frx":1B183
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Details"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   7200
         ScaleHeight     =   345
         ScaleWidth      =   2175
         TabIndex        =   32
         Top             =   5040
         Width           =   2175
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2175
            TabIndex        =   33
            Top             =   0
            Width           =   2175
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   44
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   45
                  ToolTipText     =   "Lock"
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
               TabIndex        =   42
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   43
                  ToolTipText     =   "Unlock"
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
               TabIndex        =   40
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   41
                  ToolTipText     =   "Delete"
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
               TabIndex        =   38
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  ToolTipText     =   "Refresh"
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
               TabIndex        =   36
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   4
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   37
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnContainer 
               Height          =   375
               Index           =   5
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   34
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btn 
                  Height          =   315
                  Index           =   5
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Terminate Program"
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
         Left            =   7320
         TabIndex        =   53
         Top             =   4500
         Width           =   1965
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Restriction"
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
         Left            =   5160
         TabIndex        =   52
         Top             =   4800
         Width           =   1965
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable / Disable Account"
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
         Left            =   5160
         TabIndex        =   51
         Top             =   4500
         Width           =   1965
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Level"
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
         Left            =   2880
         TabIndex        =   50
         Top             =   4800
         Width           =   2085
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete User / Records"
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
         Left            =   2880
         TabIndex        =   49
         Top             =   4500
         Width           =   2085
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit User / Records"
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
         Left            =   480
         TabIndex        =   48
         Top             =   4800
         Width           =   2205
      End
      Begin VB.Label lblRes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create User"
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
         Left            =   480
         TabIndex        =   47
         Top             =   4500
         Width           =   2205
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Restriction"
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
         Index           =   13
         Left            =   120
         TabIndex        =   46
         Top             =   4200
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Index           =   3
         Left            =   195
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Account Information"
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
         Index           =   9
         Left            =   120
         TabIndex        =   25
         Top             =   3120
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         Height          =   2000
         Index           =   2
         Left            =   195
         Top             =   3240
         Width           =   9210
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Personal Information"
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
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         Height          =   1935
         Index           =   1
         Left            =   195
         Top             =   1050
         Width           =   9210
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         Height          =   735
         Index           =   0
         Left            =   6120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status"
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
         Index           =   12
         Left            =   6840
         TabIndex        =   31
         Top             =   3480
         Width           =   2445
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "STATUS"
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
         Index           =   12
         Left            =   5880
         TabIndex        =   30
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "UserID"
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
         Index           =   11
         Left            =   1200
         TabIndex        =   29
         Top             =   3480
         Width           =   2565
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "USERID"
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
         Index           =   11
         Left            =   360
         TabIndex        =   28
         Top             =   3480
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Level"
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
         Index           =   13
         Left            =   1200
         TabIndex        =   27
         Top             =   3720
         Width           =   2565
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "LEVEL"
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
         Index           =   10
         Left            =   360
         TabIndex        =   26
         Top             =   3720
         Width           =   765
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   5880
         TabIndex        =   24
         Top             =   2205
         Width           =   885
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email"
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
         Index           =   9
         Left            =   6840
         TabIndex        =   23
         Top             =   2205
         Width           =   2445
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Email"
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
         Index           =   8
         Left            =   1080
         TabIndex        =   21
         Top             =   2205
         Width           =   2565
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   20
         Top             =   2205
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gender"
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
         Index           =   10
         Left            =   1080
         TabIndex        =   19
         Top             =   2445
         Width           =   2565
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   2445
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   1920
         Width           =   8205
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "TIME"
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
         Left            =   6240
         TabIndex        =   15
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time"
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
         Left            =   7080
         TabIndex        =   14
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
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
         Left            =   7080
         TabIndex        =   13
         Top             =   360
         Width           =   2205
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "DATE"
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
         Left            =   6240
         TabIndex        =   12
         Top             =   360
         Width           =   765
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
         Left            =   8400
         TabIndex        =   11
         Top             =   1560
         Width           =   585
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
         Index           =   0
         Left            =   6960
         TabIndex        =   10
         Top             =   1560
         Width           =   420
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
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   1560
         Width           =   495
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
         Left            =   4800
         TabIndex        =   8
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prefix"
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
         Left            =   8160
         TabIndex        =   7
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last"
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
         Left            =   6240
         TabIndex        =   6
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Middle"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   1320
         Width           =   2445
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "First"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   2565
      End
      Begin VB.Label lblname 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmAcctDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub listinfo(ID As Double)
Dim EditInfo As ADODB.Recordset
Set EditInfo = New ADODB.Recordset
If db Is Nothing Then dbconnect

With EditInfo
    .Open ("select * from T_users_acct where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    lblInfo(0).Caption = .Fields("IDnumber")
    lblInfo(1).Caption = .Fields("DateCreated")
    lblInfo(2).Caption = .Fields("TimeCreated")
    lblInfo(11).Caption = .Fields("UserID")
    lblInfo(12).Caption = .Fields("AcctStatus")
    lblInfo(13).Caption = .Fields("ulevel")
    .Close
End With
Set EditInfo = Nothing

Set EditInfo = New ADODB.Recordset
With EditInfo
    .Open ("select * from T_users_info where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    lblInfo(5).Caption = .Fields("Lname")
    lblInfo(3).Caption = .Fields("Fname")
    lblInfo(4).Caption = .Fields("Mname")
    lblInfo(6).Caption = .Fields("Prefix")
    lblInfo(7).Caption = .Fields("Address")
    lblInfo(9).Caption = .Fields("ContactNumber")
    lblInfo(8).Caption = .Fields("Email")
    lblInfo(10).Caption = .Fields("Gender")
    .Close
End With
Set EditInfo = Nothing

Set EditInfo = New ADODB.Recordset
With EditInfo
    .Open ("select * from T_acct_restric where IDnumber=" & CDbl(ID) & " "), db, 3, 3
    lblRes(0).Visible = restrictionopt(CBool(.Fields(1)))
    lblRes(1).Visible = restrictionopt(CBool(.Fields(2)))
    lblRes(2).Visible = restrictionopt(CBool(.Fields(4)))
    lblRes(3).Visible = restrictionopt(CBool(.Fields(5)))
    lblRes(4).Visible = restrictionopt(CBool(.Fields(3)))
    lblRes(5).Visible = restrictionopt(CBool(.Fields(6)))
    lblRes(6).Visible = restrictionopt(CBool(.Fields(7)))
    .Close
End With
Set EditInfo = Nothing

End Sub

Public Sub loadbutton()
btn(0).Enabled = checkPermission("btn", 3, CDbl(acctid))
btn(1).Enabled = checkPermission("btn", 3, CDbl(acctid))
btn(2).Enabled = checkPermission("btn", 4, CDbl(acctid))
btn(3).Enabled = checkPermission("btn", 2, CDbl(acctid))

If lblInfo(12).Caption = "Locked" Then
   btn(0).Enabled = False
ElseIf checkPermission("btn", 3, CDbl(acctid)) = False Then
   btn(0).Enabled = False
Else
   btn(0).Enabled = True
   btn(1).Enabled = False
End If

End Sub

Private Function restrictionopt(ResOpt As Boolean) As Integer
If ResOpt = True Then
   restrictionopt = 1
Else
   restrictionopt = 0
End If

End Function

Private Sub btn_Click(Index As Integer)
If db Is Nothing Then dbconnect
Dim checkL As ADODB.Recordset
Set checkL = New ADODB.Recordset

 checkL.Open ("select uLevel from T_users_acct where IDnumber = " & CDbl(lblInfo(0).Caption) & " "), db, 2, 2

Select Case Index
    Case 0
         If lblInfo(0).Caption = acctid Then
            MsgBox "You can't Lock your own account!   ", vbExclamation
            
            checkL.Close
            Set checkL = Nothing
         Else
             actL = "Locked"
             frmLock_Unlocl.Show 1, mdimain
         End If
         
    Case 1
         If lblInfo(0).Caption = acctid Then
            MsgBox "You can't Activate your own account!   ", vbExclamation
            
            checkL.Close
            Set checkL = Nothing
         Else
            actL = "Active"
            frmLock_Unlocl.Show 1, mdimain
         
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
               
               del.Open ("delete from T_ID_numbers WHERE IDnumber = " & CDbl(lblInfo(0).Caption) & " "), db, 2, 2
               Call UserLogs(CDbl(acctid), "Deleted ID Number [ " & lblInfo(0).Caption & " ]")
               frmUserInfo.loadinfo
               Set del = Nothing
               Set db = Nothing
               Unload Me
            End If
         End If
         
    Case 3
         If loggedname = "Manager" And checkL.Fields("Ulevel") = "Operator" Then
            edit = True
            RegFrm.loadEdit (CLng(lblInfo(0).Caption))
            RegFrm.Show 1, mdimain
            
         ElseIf loggedname = "Administrator" Then
            edit = True
            RegFrm.loadEdit (CLng(lblInfo(0).Caption))
            RegFrm.Show 1, mdimain
            
         ElseIf acctid = lblInfo(0).Caption Then
            edit = True
            RegFrm.loadEdit (CLng(lblInfo(0).Caption))
            RegFrm.Show 1, mdimain
            
         Else
            MsgBox "You dont have the right to edit this account!   ", vbExclamation
         
            checkL.Close
            Set checkL = Nothing
         
         End If
    Case 4
    
    Case 5

End Select

End Sub

Private Sub Form_Load()
If db Is Nothing Then dbconnect
listinfo (AcctInfo)
SetButtonPicture
loadbutton
countrecord

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

End Sub

Private Sub SetButtonPicture()
    With Me
        btn(0).Picture = .i16x16.ListImages(2).ExtractIcon
        btn(1).Picture = .i16x16.ListImages(3).ExtractIcon
        btn(2).Picture = .i16x16.ListImages(4).ExtractIcon
        btn(3).Picture = frmUserInfo.i16x16.ListImages(3).ExtractIcon
        btn(4).Picture = .i16x16.ListImages(1).ExtractIcon
        btn(5).Picture = .i16x16.ListImages(6).ExtractIcon

    End With
    
End Sub
