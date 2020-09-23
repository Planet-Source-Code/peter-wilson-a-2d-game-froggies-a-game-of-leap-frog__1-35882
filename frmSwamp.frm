VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSwamp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Froggies"
   ClientHeight    =   6270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "frmSwamp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSwamp.frx":030A
   ScaleHeight     =   6270
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReset 
      Caption         =   "&Replay"
      Height          =   315
      Left            =   5880
      TabIndex        =   112
      Top             =   3510
      Width           =   1245
   End
   Begin VB.TextBox txtSwampNo 
      Height          =   345
      Left            =   5805
      MaxLength       =   8
      TabIndex        =   108
      Text            =   "1"
      Top             =   1920
      Width           =   1395
   End
   Begin VB.ComboBox cmbSpeed 
      Height          =   315
      ItemData        =   "frmSwamp.frx":9A03C
      Left            =   1830
      List            =   "frmSwamp.frx":9A051
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Top             =   5730
      Width           =   1065
   End
   Begin VB.ComboBox cmbFrogColor 
      Height          =   315
      ItemData        =   "frmSwamp.frx":9A072
      Left            =   420
      List            =   "frmSwamp.frx":9A085
      Style           =   2  'Dropdown List
      TabIndex        =   105
      Top             =   5730
      Width           =   1245
   End
   Begin VB.TextBox txtFrogs 
      Height          =   285
      Left            =   1050
      TabIndex        =   103
      Text            =   "3"
      Top             =   5370
      Width           =   345
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1440
      TabIndex        =   104
      Top             =   5370
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   3
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtFrogs"
      BuddyDispid     =   196611
      OrigLeft        =   2340
      OrigTop         =   5370
      OrigRight       =   2580
      OrigBottom      =   5685
      Max             =   30
      Min             =   3
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New Swamp"
      Height          =   315
      Left            =   5880
      TabIndex        =   102
      Top             =   2790
      Width           =   1245
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   20
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   99
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   21
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   98
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   22
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   97
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   23
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   96
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   24
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   95
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   25
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   94
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   26
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   93
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   27
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   92
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   28
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   91
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   29
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   90
      Top             =   1650
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   30
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   89
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   31
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   88
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   32
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   87
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   33
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   86
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   34
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   85
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   35
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   84
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   36
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   83
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   37
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   82
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   38
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   81
      Top             =   2100
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   39
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   80
      Top             =   2100
      Width           =   455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9A0A9
            Key             =   "water"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9ABC3
            Key             =   "Red_Frog_Blink"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9B3D5
            Key             =   "grey_frog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9B6EF
            Key             =   "Blue_Frog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9C209
            Key             =   "drag_frog1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9CD23
            Key             =   "Gold_Frog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9D83D
            Key             =   "White_Frog"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9E357
            Key             =   "Purple_Frog"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9EE71
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp.frx":9F98B
            Key             =   "Red_Frog"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   99
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   79
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   98
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   78
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   97
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   77
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   96
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   76
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   95
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   75
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   94
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   74
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   93
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   73
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   92
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   72
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   91
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   71
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   90
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   70
      Top             =   4800
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   89
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   69
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   88
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   68
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   87
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   67
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   86
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   66
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   85
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   65
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   84
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   64
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   83
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   63
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   82
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   62
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   81
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   61
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   80
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   60
      Top             =   4350
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   79
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   59
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   78
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   58
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   77
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   57
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   76
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   56
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   75
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   55
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   74
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   54
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   73
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   53
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   72
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   52
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   71
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   51
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   70
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   50
      Top             =   3900
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   69
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   49
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   68
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   48
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   67
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   47
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   66
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   46
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   65
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   45
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   64
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   44
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   63
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   43
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   62
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   42
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   61
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   41
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   60
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   40
      Top             =   3450
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   59
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   39
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   58
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   38
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   57
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   37
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   56
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   36
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   55
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   35
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   54
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   34
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   53
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   33
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   52
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   32
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   51
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   31
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   50
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   30
      Top             =   3000
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   49
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   29
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   48
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   28
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   47
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   27
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   46
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   26
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   45
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   25
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   44
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   24
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   43
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   23
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   42
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   22
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   41
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   21
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   40
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   20
      Top             =   2550
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   19
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   19
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   18
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   18
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   17
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   17
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   16
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   16
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   15
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   15
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   14
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   14
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   13
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   13
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   12
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   12
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   11
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   11
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   10
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   1200
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   9
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   8
      Left            =   4050
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   8
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   7
      Left            =   3600
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   6
      Left            =   3150
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   6
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   5
      Left            =   2700
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   5
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   4
      Left            =   2250
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   4
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   3
      Left            =   1800
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   3
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   2
      Left            =   1350
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   2
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   1
      Left            =   900
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   750
      Width           =   455
   End
   Begin VB.PictureBox pictGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   455
      Index           =   0
      Left            =   450
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   750
      Width           =   455
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   6442
      TabIndex        =   113
      Top             =   930
      Width           =   120
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frogs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   5
      Left            =   450
      TabIndex        =   111
      Top             =   5400
      Width           =   570
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Score"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   0
      Left            =   5872
      TabIndex        =   110
      Top             =   660
      Width           =   1260
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Swamp No."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   3
      Left            =   5940
      TabIndex        =   109
      Top             =   1650
      Width           =   1110
   End
   Begin VB.Image imgDrag 
      Height          =   450
      Left            =   750
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Draw Speed"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1920
      TabIndex        =   107
      Top             =   5490
      Width           =   885
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "for Irene"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Index           =   2
      Left            =   5730
      TabIndex        =   101
      Top             =   4920
      Width           =   840
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Froggies !"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Index           =   1
      Left            =   5700
      TabIndex        =   100
      Top             =   4560
      Width           =   1530
   End
End
Attribute VB_Name = "frmSwamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GridSize = 10

Private Type JumpCoordinates
    FromX As Integer
    FromY As Integer
    ToX As Integer
    ToY As Integer
End Type

Dim LilleyPads(GridSize, GridSize)
Dim LilleyPadsXY(GridSize * GridSize)
Dim mColJumpHistory As Collection
Dim JumpHelper(4) As Variant

Dim frogStack As Collection


Private Sub CancelDrag()

    Dim intN As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intWhichPict As Integer
    
    
    ' Display Jump Options
    For intN = 1 To 4
        intX = JumpHelper(intN)(0)
        intY = JumpHelper(intN)(1)
        If intX <> -1 Then
            intWhichPict = (GridSize * intY) - GridSize - 1 + intX
            Me.pictGrid(intWhichPict).BorderStyle = 0
        End If
    Next intN
    
End Sub

Private Sub DisplayGrid()

    Dim intX As Integer
    Dim intY As Integer
    Dim intC As Integer
    Dim intValue As Integer
    Dim strKey As String
    Dim lngStart As Long
    Dim intN As Integer
    
    Dim intFrogCount As Integer
    
    For intY = 1 To GridSize
        For intX = 1 To GridSize
            intC = (GridSize * intY) - GridSize - 1 + intX
            
            LilleyPadsXY(intC) = Array(intX, intY)
            
            intValue = LilleyPads(intX, intY)
            Select Case intValue
                Case 1 ' Water
                    strKey = "water"
                Case 2 ' Lilley Pad
                    strKey = "leaf"
                Case 4 ' Frog
                    strKey = Me.cmbFrogColor.List(Me.cmbFrogColor.ListIndex) & "_Frog"
                    intFrogCount = intFrogCount + 1
                    
                Case 256 ' Void
                    strKey = "water"
                    
            End Select
            
            Me.pictGrid(intC).Picture = Me.ImageList1.ListImages(strKey).Picture
            Me.pictGrid(intC).Tag = intValue ' Used to speed MouseDown calculations
            
        Next intX
    Next intY
    
    If intFrogCount = 1 Then
        For intN = 1 To Val(Me.txtFrogs)
            Me.lblScore = Val(lblScore) + Val(txtFrogs) * 10
            Me.lblScore.Refresh
            Call PlaySound(App.Path & "\ribbit_short.wav", 0, 0)
        Next intN
    End If
    
    lngStart = GetTickCount
    Do
    Loop Until (GetTickCount - lngStart) > Me.cmbSpeed.ItemData(Me.cmbSpeed.ListIndex)
        
End Sub

Private Sub InitGrid()
    
    Dim intX As Integer
    Dim intY As Integer
    
    ' Create a void
    For intY = 1 To GridSize
        For intX = 1 To GridSize
            LilleyPads(intX, intY) = 256 ' Void
        Next intX
    Next intY
    
    ' Create water
    For intY = 2 To GridSize - 1
        For intX = 2 To GridSize - 1
            LilleyPads(intX, intY) = 1 ' Water
        Next intX
    Next intY
        
End Sub
Private Sub InitFrogs(HowManyFrogs As Integer)
    
    On Error GoTo errTrap
    
    Dim intCount As Integer
    Dim intRNDFrog As Integer
    Dim intDirection As Integer
    Dim intN As Integer
    Dim CurrentFrogX As Integer
    Dim CurrentFrogY As Integer
    Dim NewFrogIsJumped As Boolean
    Dim intTest1 As Integer
    Dim intTest2 As Integer
    
    Set frogStack = New Collection
        
    ' Get random sequence n, (where n = txtSwampNo.Text)
    Rnd -1
    Randomize Val(txtSwampNo.Text)
    
    ' Set initial frog position
    CurrentFrogX = Int((9 - 2 + 1) * Rnd + 2)
    CurrentFrogY = Int((9 - 2 + 1) * Rnd + 2)

    frogStack.Add Array(CurrentFrogX, CurrentFrogY)
    
    For intN = 1 To HowManyFrogs
        
        Do
            ' Get a random frog from the stack
            intRNDFrog = Int((frogStack.Count * Rnd) + 1)
            CurrentFrogX = frogStack(intRNDFrog)(0)
            CurrentFrogY = frogStack(intRNDFrog)(1)
            
            ' Choose a random jump direction (n,s,e,w)
            'intDirection = Int((4 * Rnd) + 1)
            intDirection = intDirection + 1
            If intDirection > 4 Then intDirection = 1
            Select Case intDirection
                Case 1 ' North
                    intTest1 = LilleyPads(CurrentFrogX, CurrentFrogY - 1)
                    intTest2 = LilleyPads(CurrentFrogX, CurrentFrogY - 2)
                    If (intTest2 = 2) Or (intTest2 = 1) Then ' leaf or water
                        If (intTest1 = 2) Or (intTest1 = 1) Then ' leaf or water
                        
                            ' Update the frog stack, of the frog that has done the jumping
                            frogStack.Remove intRNDFrog
                            frogStack.Add Array(CurrentFrogX, CurrentFrogY - 1)
                            frogStack.Add Array(CurrentFrogX, CurrentFrogY - 2)
                            
                            ' Save LilleyPad status (of the frog doing the jumping)
                            LilleyPads(CurrentFrogX, CurrentFrogY) = 2 ' leaf
                            LilleyPads(CurrentFrogX, CurrentFrogY - 2) = 4 ' frog
                            LilleyPads(CurrentFrogX, CurrentFrogY - 1) = 4 ' frog
                            
                            NewFrogIsJumped = True
                        End If
                    End If
                
                Case 2  ' South
                    intTest1 = LilleyPads(CurrentFrogX, CurrentFrogY + 1)
                    intTest2 = LilleyPads(CurrentFrogX, CurrentFrogY + 2)
                    If (intTest2 = 2) Or (intTest2 = 1) Then ' leaf or water
                        If (intTest1 = 2) Or (intTest1 = 1) Then ' leaf or water
                        
                            ' Update the frog stack, of the frog that has done the jumping
                            frogStack.Remove intRNDFrog
                            frogStack.Add Array(CurrentFrogX, CurrentFrogY + 1)
                            frogStack.Add Array(CurrentFrogX, CurrentFrogY + 2)
                            
                            ' Save LilleyPad status (of the frog doing the jumping)
                            LilleyPads(CurrentFrogX, CurrentFrogY) = 2 ' leaf
                            LilleyPads(CurrentFrogX, CurrentFrogY + 2) = 4 ' frog
                            LilleyPads(CurrentFrogX, CurrentFrogY + 1) = 4 ' frog
                            
                            NewFrogIsJumped = True
                        End If
                    End If

                Case 3 ' East
                    intTest1 = LilleyPads(CurrentFrogX + 1, CurrentFrogY)
                    intTest2 = LilleyPads(CurrentFrogX + 2, CurrentFrogY)
                    If (intTest2 = 2) Or (intTest2 = 1) Then ' leaf or water
                        If (intTest1 = 2) Or (intTest1 = 1) Then ' leaf or water
                        
                            ' Update the frog stack, of the frog that has done the jumping
                            frogStack.Remove intRNDFrog
                            frogStack.Add Array(CurrentFrogX + 1, CurrentFrogY)
                            frogStack.Add Array(CurrentFrogX + 2, CurrentFrogY)
                            
                            ' Save LilleyPad status (of the frog doing the jumping)
                            LilleyPads(CurrentFrogX, CurrentFrogY) = 2 ' leaf
                            LilleyPads(CurrentFrogX + 1, CurrentFrogY) = 4  ' frog
                            LilleyPads(CurrentFrogX + 2, CurrentFrogY) = 4 ' frog
                            
                            NewFrogIsJumped = True
                        End If
                    End If
                    
                Case 4 ' West
                    intTest1 = LilleyPads(CurrentFrogX - 1, CurrentFrogY)
                    intTest2 = LilleyPads(CurrentFrogX - 2, CurrentFrogY)
                    If (intTest2 = 2) Or (intTest2 = 1) Then ' leaf or water
                        If (intTest1 = 2) Or (intTest1 = 1) Then ' leaf or water
                        
                            ' Update the frog stack, of the frog that has done the jumping
                            frogStack.Remove intRNDFrog
                            frogStack.Add Array(CurrentFrogX - 1, CurrentFrogY)
                            frogStack.Add Array(CurrentFrogX - 2, CurrentFrogY)
                            
                            ' Save LilleyPad status (of the frog doing the jumping)
                            LilleyPads(CurrentFrogX, CurrentFrogY) = 2 ' leaf
                            LilleyPads(CurrentFrogX - 1, CurrentFrogY) = 4  ' frog
                            LilleyPads(CurrentFrogX - 2, CurrentFrogY) = 4 ' frog
                            
                            NewFrogIsJumped = True
                        End If
                    End If
                
            End Select
        
resumePoint:
        
        Loop Until NewFrogIsJumped = True
        NewFrogIsJumped = False
        
        If Me.cmbSpeed.ItemData(Me.cmbSpeed.ListIndex) <> 0 Then Call DisplayGrid

    Next intN
    
    'Call PlaySound("F:\Multimedia\wave files\frog2[1].wav", 0, SND_ASYNC + SND_NOSTOP)
    
       
    If Me.cmbSpeed.ItemData(Me.cmbSpeed.ListIndex) = 0 Then Call DisplayGrid
    
    Exit Sub
errTrap:
    Resume resumePoint
    
End Sub

Private Sub JumpFrog(varJumpPos As Variant)

    Dim intFromX As Integer
    Dim intFromY As Integer
    Dim intToX As Integer
    Dim intToY As Integer
    
    Dim intMidFrogX As Integer
    Dim intMidFrogY As Integer
    
    ' Add new jump to the first position on the stack.
    mColJumpHistory.Add varJumpPos
    
    ' From (old frog position)
    intFromX = varJumpPos(0)
    intFromY = varJumpPos(1)
    LilleyPads(intFromX, intFromY) = 2 ' lilley pad
    
    ' To (new frog position)
    intToX = varJumpPos(2)
    intToY = varJumpPos(3)
    LilleyPads(intToX, intToY) = 4 ' frog
    
    ' Erase the frog that was jumped over.
    intMidFrogX = (intFromX + intToX) / 2
    intMidFrogY = (intFromY + intToY) / 2
    LilleyPads(intMidFrogX, intMidFrogY) = 2 ' lilley pad
    
    Call DisplayGrid
    
End Sub

Public Sub RemoveFrogFromStack(X As Integer, Y As Integer)

    Dim intN As Integer
    
    For intN = 1 To frogStack.Count
        If (frogStack(intN)(0) = X) And (frogStack(intN)(1) = Y) Then
            frogStack.Remove intN
            Exit For
        End If
    Next intN
    
End Sub


Private Sub btnNew_Click()

    ' Create Random Swamp  number
    Me.txtSwampNo = Int((2147000000 * Rnd) + 1)
    
    txtFrogs = Val(txtFrogs) + 1
    If Val(txtFrogs) > 30 Then txtFrogs = 30
    
    Call btnReset_Click

End Sub

Private Sub btnReset_Click()

    Screen.MousePointer = vbHourglass
    Me.btnNew.Enabled = False
    
    ' Reset Jump History
    Set mColJumpHistory = Nothing
    Set mColJumpHistory = New Collection
    
    Call InitGrid
    Call InitFrogs(txtFrogs - 1)
    
    Me.btnNew.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub cmbSpeed_Click()

    If cmbSpeed.ListIndex < 2 Then
        If Val(txtFrogs) > 7 Then txtFrogs = 7
    End If
    
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

    Call CancelDrag

End Sub

Private Sub Form_Load()
    
    Me.Show
    
    Me.cmbFrogColor.ListIndex = 0
    Me.cmbSpeed.ListIndex = 3
    
    Call btnNew_Click

End Sub

Private Sub pictGrid_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
    Dim intN As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim FromTo As JumpCoordinates
    Dim varJumpPos As Variant
    
    ' To Coordinates
    intX = LilleyPadsXY(Index)(0)
    intY = LilleyPadsXY(Index)(1)
    
    For intN = 1 To 4
        If intX = JumpHelper(intN)(0) Then
            If intY = JumpHelper(intN)(1) Then
                ' The Frog is allowed to jump here.
                
                ' Store the From/To values into an array
                varJumpPos = Array(LilleyPadsXY(Source.Tag)(0), LilleyPadsXY(Source.Tag)(1), intX, intY)
                                
                ' Update Display and Jump history
                Call JumpFrog(varJumpPos)
                
                Call CancelDrag
                ' Do Stuff here!
                Exit Sub
            End If
        End If
    Next intN

    ' An invalid drop has occured
    Call CancelDrag
    
End Sub

Private Sub pictGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim intIsFrog As Integer
    Dim intIsLilleyPad As Integer
    Dim blnFrogCanJump As Boolean
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intN As Integer
    Dim intWhichPict As Integer
    
    
    ' Check if a Frog has been clicked
    If pictGrid(Index).Tag <> 4 Then Exit Sub
    
    ' Convert the current Index number into the X Y co-ordinate system
    intX = LilleyPadsXY(Index)(0)
    intY = LilleyPadsXY(Index)(1)
    
    ' Can this frog jump another frog, and land on an empty lilley pad?
    blnFrogCanJump = False
    For intN = 1 To 4
        JumpHelper(intN) = Array(-1, -1)
    Next intN
    
    ' Check the North
    intIsFrog = LilleyPads(intX, intY - 1)
    If intIsFrog = 4 Then ' There is a frog north
        intIsLilleyPad = LilleyPads(intX, intY - 2)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            blnFrogCanJump = True
            JumpHelper(1) = Array(intX, intY - 2)
        End If
    End If
    
    ' Check the South
    intIsFrog = LilleyPads(intX, intY + 1)
    If intIsFrog = 4 Then ' There is a frog north
        intIsLilleyPad = LilleyPads(intX, intY + 2)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            blnFrogCanJump = True
            JumpHelper(2) = Array(intX, intY + 2)
        End If
    End If
    
    ' Check the East
    intIsFrog = LilleyPads(intX + 1, intY)
    If intIsFrog = 4 Then ' There is a frog north
        intIsLilleyPad = LilleyPads(intX + 2, intY)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            blnFrogCanJump = True
            JumpHelper(3) = Array(intX + 2, intY)
        End If
    End If
    
    ' Check the West
    intIsFrog = LilleyPads(intX - 1, intY)
    If intIsFrog = 4 Then ' There is a frog north
        intIsLilleyPad = LilleyPads(intX - 2, intY)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            blnFrogCanJump = True
            JumpHelper(4) = Array(intX - 2, intY)
        End If
    End If
    
    If blnFrogCanJump = False Then
        Screen.MousePointer = vbNoDrop
            pictGrid(Index).Picture = Me.ImageList1.ListImages("Red_Frog_Blink").Picture
            Call PlaySound(App.Path & "\ribbit_short.wav", 0, 0)
            pictGrid(Index).Picture = Me.ImageList1.ListImages("Red_Frog").Picture
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
FrogCanJump:
    
    ' Display Jump Options
    For intN = 1 To 4
        intX = JumpHelper(intN)(0)
        intY = JumpHelper(intN)(1)
        If intX <> -1 Then
            intWhichPict = (GridSize * intY) - GridSize - 1 + intX
            Me.pictGrid(intWhichPict).BorderStyle = 1
        End If
    Next intN
    
    ' Enable this picture for dragging
    If pictGrid(Index).Tag = 4 Then
        imgDrag.Tag = Index ' Set this so the Drop/DragOver event knows the source index
        imgDrag.Move pictGrid(Index).Left + X, pictGrid(Index).Top + Y
        imgDrag.DragIcon = Me.ImageList1.ListImages("drag_frog1").ExtractIcon
        imgDrag.Drag  ' Drag outline.
    End If
    
End Sub

Private Sub txtFrogs_Change()

    txtFrogs = Val(txtFrogs)
    If txtFrogs > 30 Then txtFrogs = 30
    
    If txtFrogs > 7 Then
        If Me.cmbSpeed.ListIndex < 3 Then
            Me.cmbSpeed.ListIndex = 2
        End If
    End If
    
End Sub

