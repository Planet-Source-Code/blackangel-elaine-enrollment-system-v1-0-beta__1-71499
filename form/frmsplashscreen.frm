VERSION 5.00
Begin VB.Form frmsplashscreen 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Palette         =   "frmsplashscreen.frx":0000
   Picture         =   "frmsplashscreen.frx":1BCB8
   ScaleHeight     =   3000
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   360
      Top             =   2400
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "(0%)..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmsplashscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintCount As Integer, mintPause As Integer
Private Sub Timer1_Timer()
    mintPause = mintPause + 1
    If mintCount < 100 Then
        mintCount = mintCount + 2
        lblCount.Caption = "(" & mintCount & "%)..."
      End If
        If mintPause = 100 Then
        lblCount.Caption = "App..."
        lblInform.Caption = "Starting"
    ElseIf mintPause = 110 Then
        Unload Me
       frmchoose_login.Show vbModal
    End If
  End Sub


