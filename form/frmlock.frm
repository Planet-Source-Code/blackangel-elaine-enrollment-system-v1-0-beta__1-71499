VERSION 5.00
Begin VB.Form frmlock 
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1720
      FrameColor      =   65280
      Caption         =   "Locked:"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16777152
      ColorTo         =   12648384
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the password:"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Me.txtPassword = current_info.password Then
Unload Me
Else
End If
End If

End Sub
