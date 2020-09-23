VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2580
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   4800
      Top             =   420
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.midar.com.au"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "www.midar.com"
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2002 - MIDAR Pty Ltd"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "www.midar.com"
      Top             =   2520
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   60
      Picture         =   "frmSplash.frx":0442
      Top             =   30
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   5340
      X2              =   -60
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Image imgPuzzles 
      Height          =   300
      Index           =   1
      Left            =   4080
      Picture         =   "frmSplash.frx":0F74
      Top             =   0
      Width           =   300
   End
   Begin VB.Image imgPuzzles 
      Height          =   300
      Index           =   3
      Left            =   4920
      Picture         =   "frmSplash.frx":1466
      Top             =   0
      Width           =   300
   End
   Begin VB.Image imgPuzzles 
      Height          =   300
      Index           =   2
      Left            =   4500
      Picture         =   "frmSplash.frx":1958
      Top             =   0
      Width           =   300
   End
   Begin VB.Image imgPuzzles 
      Height          =   300
      Index           =   0
      Left            =   3675
      Picture         =   "frmSplash.frx":1E4A
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Professional Leap-Frog Game"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   60
      Width           =   2085
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MIDAR's Froggies Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   2070
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   300
      Index           =   2
      Left            =   -1140
      Top             =   0
      Width           =   5115
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   0
      Picture         =   "frmSplash.frx":233C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2475
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"37D47C2D01C2"
Option Explicit

Private Sub btnClose_Click()

    Unload Me
    
End Sub

Private Sub Form_Click()

    If Me.btnClose.Visible = False Then Unload Me
    
End Sub

Private Sub Form_Load()

    Me.lblAbout(0).Caption = AppTitle
    Me.lblAbout(1).Caption = "Copyright © " & Year(Now) & " - MIDAR Pty Ltd"
    
End Sub

Private Sub Timer1_Timer()

    Unload Me
    
End Sub


