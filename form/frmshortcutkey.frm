VERSION 5.00
Begin VB.Form frmshortcutkey 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shortcut Keys"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   0
   End
   Begin EnrollmentSystem.ChameleonBtn cmdok 
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15724527
      BCOLO           =   15724527
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmshortcutkey.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Shortcut keys"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         Height          =   495
         Left            =   2400
         TabIndex        =   16
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Shortcut Key"
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl +  N"
         Height          =   495
         Left            =   2400
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   " Notepad"
         Height          =   495
         Left            =   960
         TabIndex        =   13
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl +  c"
         Height          =   495
         Left            =   2400
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Calculator"
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "System info"
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl +  X"
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CLOSE"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl +  o"
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Logout"
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl +  L"
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock the System"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmshortcutkey.frx":001C
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmshortcutkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bhide As Boolean

Private Sub cmdok_Click()
Timer1.Interval = 30
bhide = True
End Sub

Private Sub Form_Activate()
Me.Height = 0
bhide = False
Timer1.Interval = 30
End Sub



Private Sub Timer1_Timer()
If bhide = False Then
    If Me.Height >= 3765 Then
        Me.Width = 4080
        Me.Height = 3765
        Timer1.Interval = 0
        Else
        Me.Height = Me.Height + 300
    End If
Else
 If Me.Height <= 600 Then
        Me.Width = 0
        Me.Height = 0
        Timer1.Interval = 0
        Unload Me
        Else
        Me.Height = Me.Height - 300
        DoEvents
    End If
End If

 Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
