VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSwamp2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Froggies"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7500
   ClipControls    =   0   'False
   Icon            =   "frmSwamp2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSwamp2.frx":030A
   ScaleHeight     =   5715
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6870
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnUndo 
      Caption         =   "&Undo"
      Height          =   315
      Left            =   5888
      TabIndex        =   9
      ToolTipText     =   "Undo"
      Top             =   3930
      Width           =   1245
   End
   Begin VB.CheckBox chkHelper 
      Height          =   195
      Left            =   5790
      TabIndex        =   7
      Top             =   4830
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "&Replay"
      Height          =   315
      Left            =   5888
      TabIndex        =   6
      ToolTipText     =   "Replay the current swamp"
      Top             =   3570
      Width           =   1245
   End
   Begin VB.TextBox txtSwampNo 
      Height          =   345
      Left            =   5813
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "1"
      Top             =   1980
      Width           =   1395
   End
   Begin VB.TextBox txtFrogs 
      Height          =   285
      Left            =   6420
      TabIndex        =   1
      Text            =   "2"
      Top             =   4530
      Width           =   345
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   6810
      TabIndex        =   2
      Top             =   4530
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   3
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtFrogs"
      BuddyDispid     =   196613
      OrigLeft        =   2340
      OrigTop         =   5370
      OrigRight       =   2580
      OrigBottom      =   5685
      Max             =   60
      Min             =   3
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New Swamp"
      Height          =   315
      Left            =   5888
      TabIndex        =   0
      ToolTipText     =   "Starts a new Swamp"
      Top             =   2850
      Width           =   1245
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp2.frx":71D3
            Key             =   "blink"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp2.frx":7E25
            Key             =   "drag"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp2.frx":8A77
            Key             =   "frog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp2.frx":96C9
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSwamp2.frx":A31B
            Key             =   "water"
         EndProperty
      EndProperty
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Moves Left - Try Again!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   1365
      TabIndex        =   14
      Top             =   180
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moves Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   5940
      TabIndex        =   13
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblMovesLeft 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6435
      TabIndex        =   12
      Top             =   1290
      Width           =   150
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   5820
      TabIndex        =   11
      Top             =   420
      Width           =   1380
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6435
      TabIndex        =   10
      Top             =   690
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Free Lilleys"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6030
      TabIndex        =   8
      Top             =   4830
      Width           =   1230
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   99
      Left            =   4620
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   98
      Left            =   4140
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   97
      Left            =   3660
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   96
      Left            =   3180
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   95
      Left            =   2700
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   94
      Left            =   2220
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   93
      Left            =   1740
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   92
      Left            =   1260
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   91
      Left            =   780
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   90
      Left            =   300
      Top             =   5010
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   89
      Left            =   4620
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   88
      Left            =   4140
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   87
      Left            =   3660
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   86
      Left            =   3180
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   85
      Left            =   2700
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   84
      Left            =   2220
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   83
      Left            =   1740
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   82
      Left            =   1260
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   81
      Left            =   780
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   80
      Left            =   300
      Top             =   4530
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   79
      Left            =   4620
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   78
      Left            =   4140
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   77
      Left            =   3660
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   76
      Left            =   3180
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   75
      Left            =   2700
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   74
      Left            =   2220
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   73
      Left            =   1740
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   72
      Left            =   1260
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   71
      Left            =   780
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   70
      Left            =   300
      Top             =   4050
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   69
      Left            =   4620
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   68
      Left            =   4140
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   67
      Left            =   3660
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   66
      Left            =   3180
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   65
      Left            =   2700
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   64
      Left            =   2220
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   63
      Left            =   1740
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   62
      Left            =   1260
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   61
      Left            =   780
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   60
      Left            =   300
      Top             =   3570
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   59
      Left            =   4620
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   58
      Left            =   4140
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   57
      Left            =   3660
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   56
      Left            =   3180
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   55
      Left            =   2700
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   54
      Left            =   2220
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   53
      Left            =   1740
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   52
      Left            =   1260
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   51
      Left            =   780
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   50
      Left            =   300
      Top             =   3090
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   49
      Left            =   4620
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   48
      Left            =   4140
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   47
      Left            =   3660
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   46
      Left            =   3180
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   45
      Left            =   2700
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   44
      Left            =   2220
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   43
      Left            =   1740
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   42
      Left            =   1260
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   41
      Left            =   780
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   40
      Left            =   300
      Top             =   2610
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   39
      Left            =   4620
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   38
      Left            =   4140
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   37
      Left            =   3660
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   36
      Left            =   3180
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   35
      Left            =   2700
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   34
      Left            =   2220
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   33
      Left            =   1740
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   32
      Left            =   1260
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   31
      Left            =   780
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   30
      Left            =   300
      Top             =   2130
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   29
      Left            =   4620
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   28
      Left            =   4140
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   27
      Left            =   3660
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   26
      Left            =   3180
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   25
      Left            =   2700
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   24
      Left            =   2220
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   23
      Left            =   1740
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   22
      Left            =   1260
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   21
      Left            =   780
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   20
      Left            =   300
      Top             =   1650
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   19
      Left            =   4620
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   18
      Left            =   4140
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   17
      Left            =   3660
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   16
      Left            =   3180
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   15
      Left            =   2700
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   14
      Left            =   2220
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   13
      Left            =   1740
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   12
      Left            =   1260
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   11
      Left            =   780
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   10
      Left            =   300
      Top             =   1170
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   9
      Left            =   4620
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   8
      Left            =   4140
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   7
      Left            =   3660
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   6
      Left            =   3180
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   5
      Left            =   2700
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   4
      Left            =   2220
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   3
      Left            =   1740
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   2
      Left            =   1260
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   1
      Left            =   780
      Top             =   690
      Width           =   480
   End
   Begin VB.Image imgGrid 
      Height          =   480
      Index           =   0
      Left            =   300
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frogs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Index           =   5
      Left            =   5790
      TabIndex        =   5
      Top             =   4560
      Width           =   630
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Swamp No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   5910
      TabIndex        =   4
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Image imgDrag 
      Height          =   450
      Left            =   660
      Top             =   120
      Width           =   450
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFItem 
         Caption         =   "&Change Player's Details..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFItem 
         Caption         =   "-"
         Index           =   98
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFItem 
         Caption         =   "&Quit Game"
         Index           =   99
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHItem 
         Caption         =   "Contents..."
         Index           =   0
      End
      Begin VB.Menu mnuHItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuHItem 
         Caption         =   "About this Game..."
         Index           =   99
      End
   End
End
Attribute VB_Name = "frmSwamp2"
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

' Image Keys
Dim m_strWaterKey As String
Dim m_strLeafKey As String
Dim m_strFrogKey As String
Dim m_strBlinkKey As String
Dim m_strDragKey As String
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
            If chkHelper.Value = vbChecked Then
                imgGrid(intWhichPict).BorderStyle = 0
            End If
        End If
    Next intN
    
End Sub

Private Function CanFrogJump(ByVal Index As Integer) As Boolean
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intN As Integer
    Dim intIsFrog As Integer
    Dim intIsLilleyPad As Integer
    
    CanFrogJump = False
    
    ' Convert the current Index number into the X Y co-ordinate system
    intX = LilleyPadsXY(Index)(0)
    intY = LilleyPadsXY(Index)(1)
    
    ' Can this frog jump another frog, and land on an empty lilley pad?
    CanFrogJump = False
    For intN = 1 To 4
        JumpHelper(intN) = Array(-1, -1)
    Next intN
    
    ' Check the North
    intIsFrog = LilleyPads(intX, intY - 1)
    If intIsFrog = 4 Then ' There is a frog North
        intIsLilleyPad = LilleyPads(intX, intY - 2)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            CanFrogJump = True
            JumpHelper(1) = Array(intX, intY - 2)
        End If
    End If
    
    ' Check the South
    intIsFrog = LilleyPads(intX, intY + 1)
    If intIsFrog = 4 Then ' There is a frog South
        intIsLilleyPad = LilleyPads(intX, intY + 2)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            CanFrogJump = True
            JumpHelper(2) = Array(intX, intY + 2)
        End If
    End If
    
    ' Check the East
    intIsFrog = LilleyPads(intX + 1, intY)
    If intIsFrog = 4 Then ' There is a frog East
        intIsLilleyPad = LilleyPads(intX + 2, intY)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            CanFrogJump = True
            JumpHelper(3) = Array(intX + 2, intY)
        End If
    End If
    
    ' Check the West
    intIsFrog = LilleyPads(intX - 1, intY)
    If intIsFrog = 4 Then ' There is a frog West
        intIsLilleyPad = LilleyPads(intX - 2, intY)
        If intIsLilleyPad = 2 Then ' There is a lilley pad
            CanFrogJump = True
            JumpHelper(4) = Array(intX - 2, intY)
        End If
    End If

End Function

Private Function CanFrogsMove() As Boolean

    Dim varC As Variant
    Dim intMovesLeft As Integer
    
    CanFrogsMove = False
    
    intMovesLeft = 0
    For Each varC In frogStack
        CanFrogsMove = CanFrogJump(CInt(varC))
        If CanFrogsMove = True Then intMovesLeft = intMovesLeft + 1
    Next varC

    If intMovesLeft > 0 Then CanFrogsMove = True
    Me.lblMovesLeft.Caption = intMovesLeft
    
End Function

Private Sub DisplayGrid()

    On Error GoTo errTrap
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intC As Integer
    Dim intValue As Integer
    Dim strKey As String
    Dim lngStart As Long
    Dim intN As Integer
    Dim varC As Variant
    
    Set frogStack = Nothing
    Set frogStack = New Collection
    
    For intY = 1 To GridSize
        For intX = 1 To GridSize
            
            intC = (GridSize * intY) - GridSize - 1 + intX
            intValue = LilleyPads(intX, intY)
            LilleyPadsXY(intC) = Array(intX, intY)
            
            Select Case intValue
                Case 1 ' Water
                    strKey = m_strWaterKey
                Case 2 ' Lilley Pad
                    strKey = m_strLeafKey
                Case 4 ' Frog
                    strKey = m_strFrogKey
                    frogStack.Add intC
                    
                Case 256 ' Void
                    strKey = m_strWaterKey
                    
            End Select
            
            If Val(imgGrid(intC).Tag) <> intValue Then
                Me.imgGrid(intC).Picture = Me.ImageList1.ListImages(strKey).ExtractIcon
                Me.imgGrid(intC).Tag = intValue ' Used to speed MouseDown calculations
            End If
            
        Next intX
    Next intY
    
    ' Check for End of Swamp
    If frogStack.Count = 1 Then
        Me.Refresh
        For intN = 1 To Val(Me.txtFrogs)
            Me.lblScore = Val(lblScore) + 10
            Me.lblScore.Refresh
            Call PlaySound(App.Path & "\blink.wav", 0, SND_ASYNC)
            Call PauseFor(120)
        Next intN
            
        Call btnNew_Click
        'Exit Sub
    End If

    ' Are the Frogs stuck with no more moves?
    If CanFrogsMove = False Then
        Me.Refresh
        Me.Labels(6).Visible = True
        Me.Labels(4).Caption = "Try Again"
        Call PauseFor(2000)
        For Each varC In frogStack
            imgGrid(CInt(varC)).Picture = Me.ImageList1.ListImages(m_strBlinkKey).ExtractIcon
            imgGrid(CInt(varC)).Refresh
            Call PlaySound(App.Path & "\blink.wav", 0, 0)
            imgGrid(CInt(varC)).Picture = Me.ImageList1.ListImages(m_strFrogKey).ExtractIcon
        Next varC
                
        Me.Labels(4).Caption = "Moves Left"
        Me.Labels(6).Visible = False
        Call btnReset_Click
        Exit Sub
    End If
    
    Exit Sub
errTrap:
    
End Sub

Private Sub InitGrid()
    
    ' This section kind-of draws an invisible boundry around the edge of the screen
    ' so that Frogs can't move off screen.  Actually, most of the game intelligence
    ' has absolutely nothing to do with the graphics - it's all done using arrays
    ' in memory.  When everything is calculated, we just update the display with
    ' whatever is in the array.
    
    Dim intX As Integer
    Dim intY As Integer
    
    ' Create a void (in memory)
    For intY = 1 To GridSize
        For intX = 1 To GridSize
            LilleyPads(intX, intY) = 256 ' Void are where Frogs can jump too.
        Next intX
    Next intY
    
    ' Create water
    For intY = 2 To GridSize - 1
        For intX = 2 To GridSize - 1
            LilleyPads(intX, intY) = 1 ' Water, frogs can't move here either.
        Next intX
    Next intY
        
End Sub
Private Sub InitFrogs(HowManyFrogs As Integer)
    
    ' This routine creates a new swamp, by figuring out where all the frogs and
    ' lilley pads will go.
    ' Here's how it's done:
    ' STEP 1) Push a single frog onto the stack (temporary storage area)
    '         BEGIN LOOP
    ' STEP 2) Choose a random frog from the stack (at first there's just one
    '         frog, but there will be more with each pass of the loop)
    ' STEP 3) Choose a random jump location and test to see if the frog can
    '         jump there.
    ' STEP 4) If the frog was able to jump there, then add a lilley at the appropiate
    '         place, then add the frog to the new location, then add this new frog
    '         to the stack.

    
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
    Dim intFailedAttempts As Integer
    
    Set frogStack = New Collection
        
    ' Get random sequence n, (where n = txtSwampNo.Text)
    Rnd -1
    Randomize Val(txtSwampNo.Text)
    
    ' Set initial frog position (STEP 1)
    CurrentFrogX = Int((9 - 2 + 1) * Rnd + 2)
    CurrentFrogY = Int((9 - 2 + 1) * Rnd + 2)
    frogStack.Add Array(CurrentFrogX, CurrentFrogY)
    
    
    For intN = 1 To HowManyFrogs
        
        intFailedAttempts = -1
        Do
            intFailedAttempts = intFailedAttempts + 1
            
            ' Get a random frog from the stack
            intRNDFrog = Int((frogStack.Count * Rnd) + 1)
            CurrentFrogX = frogStack(intRNDFrog)(0)
            CurrentFrogY = frogStack(intRNDFrog)(1)
            
            ' Choose a random jump direction (n,s,e,w)
            intDirection = Int((4 * Rnd) + 1)

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

        ' Try to add a new frog 6 times before giving up.
        ' (Usually, it gives up when the board is full and there's nowhere left
        ' for new frogs to go)
        Loop Until (NewFrogIsJumped = True) Or (intFailedAttempts > 6)
        NewFrogIsJumped = False

    Next intN
    
    ' Update the display when finished.
    Call DisplayGrid
    
    Exit Sub
errTrap:
    If Err.Number = 9 Then Resume resumePoint
    
End Sub

Private Sub JumpFrog(varJumpPos As Variant)

    Dim intFromX As Integer
    Dim intFromY As Integer
    Dim intToX As Integer
    Dim intToY As Integer
    
    Dim intMidFrogX As Integer
    Dim intMidFrogY As Integer
    
    ' Add new jump to the first position on the stack.
    If mColJumpHistory.Count = 0 Then
        mColJumpHistory.Add varJumpPos
    Else
        mColJumpHistory.Add varJumpPos, , 1
    End If
    
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

Private Sub LoadBitmaps()

    ' PLEASE - PLEASE - PLEASE design your own pictures, it would be really cool
    ' to see a game of "Alien Jump", or "Croc Jumper".... or whatever!
    ' Someone out there must be good at designing graphics!
        
    On Error GoTo errTrap
    
    Dim strFilePath As String
    
    ' Set Defaults in case of error
    m_strWaterKey = "water"
    m_strLeafKey = "leaf"
    m_strFrogKey = "frog"
    m_strBlinkKey = "blink"
    m_strDragKey = "drag"
    
    strFilePath = App.Path & "\bitmaps\"
    
    ' Load Background
    Me.Picture = LoadPicture(strFilePath & "backgnd.jpg")
    
    ' Load Water
    Call Me.ImageList1.ListImages.Add(, "water1", LoadPicture(strFilePath & "water.bmp"))
    m_strWaterKey = "water1"
    
    ' Load Leaf
    Call Me.ImageList1.ListImages.Add(, "leaf1", LoadPicture(strFilePath & "leaf.bmp"))
    m_strLeafKey = "leaf1"
    
    ' Load Frog
    Call Me.ImageList1.ListImages.Add(, "frog1", LoadPicture(strFilePath & "frog.bmp"))
    m_strFrogKey = "frog1"
    
    ' Load Blinking Frog
    Call Me.ImageList1.ListImages.Add(, "blink1", LoadPicture(strFilePath & "blink.bmp"))
    m_strBlinkKey = "blink1"
    
    ' Load Drag Image
    Call Me.ImageList1.ListImages.Add(, "drag1", LoadPicture(strFilePath & "drag.bmp"))
    m_strDragKey = "drag1"
    
    
    Exit Sub
errTrap:

End Sub

Private Sub PauseFor(milliSec As Long)

    Dim lngStart As Long
    
    lngStart = GetTickCount
    Do
    Loop Until (GetTickCount - lngStart) > milliSec


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
    
    ' Set number of frogs
    txtFrogs = Val(txtFrogs) + 1
    If Val(txtFrogs) > Me.UpDown1.Max Then txtFrogs = Me.UpDown1.Max
    
    ' Penalize for new game
    lblScore = lblScore - 10
    
    ' Reset the game
    Call btnReset_Click
    
End Sub

Private Sub btnReset_Click()

    Screen.MousePointer = vbHourglass
    Me.btnNew.Enabled = False
    
    ' Reset Jump History
    Set mColJumpHistory = Nothing
    Set mColJumpHistory = New Collection
    
    Call InitGrid
    Call InitFrogs(txtFrogs)
    
    Me.btnNew.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub btnUndo_Click()

    ' Basically the reverse of JumpFrog
    ' All Jump moves are stored into a stack, so we can easily Undo, or Redo.
    
    Dim varJumpPos As Variant
    Dim intFromX As Integer
    Dim intFromY As Integer
    Dim intToX As Integer
    Dim intToY As Integer
    
    Dim intMidFrogX As Integer
    Dim intMidFrogY As Integer
    
    If mColJumpHistory.Count = 0 Then Exit Sub
    
    ' Remove the last move on the stack
    varJumpPos = mColJumpHistory(1)
    mColJumpHistory.Remove 1
    
    ' From
    intFromX = varJumpPos(2)
    intFromY = varJumpPos(3)
    LilleyPads(intFromX, intFromY) = 2 ' lilley pad
    
    ' To
    intToX = varJumpPos(0)
    intToY = varJumpPos(1)
    LilleyPads(intToX, intToY) = 4 ' frog
    
    ' Erase the frog that was jumped over.
    intMidFrogX = (intFromX + intToX) / 2
    intMidFrogY = (intFromY + intToY) / 2
    LilleyPads(intMidFrogX, intMidFrogY) = 4 ' frog
    
    ' Penalize Score
    lblScore = lblScore - 20
    
    Call DisplayGrid
    
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

    Call CancelDrag

End Sub

Private Sub Form_Load()
        
    Me.Show
    
    ' You can change the pictures..
    ' Try designing your own birds, crocodiles, aliens, bugs, whatever!!!
    ' Oh... and don't forget to change the background to match!
    Call LoadBitmaps
        
    frmSplash.Show vbModeless, Me
    
    ' Simulate the user pressing on the New Button
    Call btnNew_Click
    
    ' Load previous score
    lblScore.Caption = Val(GetSetting(AppTitle, "UserDetails", "Score", ""))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Save the user's score in the Registry.
    Call SaveSetting(AppTitle, "UserDetails", "Score", lblScore.Caption)
    
End Sub

Private Sub imgGrid_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
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
                Call CancelDrag
                
                ' Store the From/To values into an array
                varJumpPos = Array(LilleyPadsXY(Source.Tag)(0), LilleyPadsXY(Source.Tag)(1), intX, intY)
                                
                ' Update Display and Jump history
                Call JumpFrog(varJumpPos)
                
                Exit Sub
            End If
        End If
    Next intN

    ' An invalid drop has occured
    Call CancelDrag
    
End Sub

Private Sub imgGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim intIsFrog As Integer
    Dim intIsLilleyPad As Integer
    Dim blnFrogCanJump As Boolean
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intN As Integer
    Dim intWhichPict As Integer
    
    ' Check if a Frog has been clicked
    If imgGrid(Index).Tag <> 4 Then Exit Sub
    
    If CanFrogJump(Index) = False Then
        ' If the Frog can NOT jump, croak and blink his eyes.
        Screen.MousePointer = vbNoDrop
            imgGrid(Index).Picture = Me.ImageList1.ListImages(m_strBlinkKey).ExtractIcon
            imgGrid(Index).Refresh
            Call PlaySound(App.Path & "\blink.wav", 0, 0)
            imgGrid(Index).Picture = Me.ImageList1.ListImages(m_strFrogKey).ExtractIcon
        Screen.MousePointer = vbDefault
    Else
        ' Frog CAN jump
        
        ' Display Jump Options
        If chkHelper.Value = vbChecked Then
            For intN = 1 To 4
                intX = JumpHelper(intN)(0)
                intY = JumpHelper(intN)(1)
                If intX <> -1 Then
                    intWhichPict = (GridSize * intY) - GridSize - 1 + intX
                    imgGrid(intWhichPict).BorderStyle = 1
                End If
            Next intN
        End If
        
        ' Enable this picture for dragging
        imgDrag.Tag = Index ' Set this so the Drop/DragOver event knows the source index
        imgDrag.Move imgGrid(Index).Left + X, imgGrid(Index).Top + Y
        imgDrag.DragIcon = Me.ImageList1.ListImages(m_strDragKey).ExtractIcon
        imgDrag.Drag  ' Drag outline.
    End If
    
End Sub

Private Sub Labels_DblClick(Index As Integer)

    Dim intA As Integer
    
    intA = MsgBox("Reset score to zero?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Clear Score?")
    If intA = vbYes Then lblScore.Caption = 0
    
End Sub


Private Sub mnuFItem_Click(Index As Integer)

    Dim strTemp As String
    
    Select Case Index
        Case 0 ' Change Player's Details
            strTemp = InputBox("What is the name of the player?", "Change Player Name", g_DisplayName)
            If strTemp <> "" Then g_DisplayName = strTemp
            
        Case Else
            Unload Me
            
    End Select
    
End Sub

Private Sub mnuHItem_Click(Index As Integer)

    Select Case Index
        Case 0 ' Contents
            Me.CommonDialog1.HelpFile = App.Path & "/froggies.hlp"
            Me.CommonDialog1.HelpCommand = &HB
            Me.CommonDialog1.ShowHelp
            
        Case 99 ' About
            frmSplash.btnClose.Visible = True
            frmSplash.Timer1.Enabled = False
            frmSplash.Show vbModeless, Me
    
    End Select
    
End Sub

