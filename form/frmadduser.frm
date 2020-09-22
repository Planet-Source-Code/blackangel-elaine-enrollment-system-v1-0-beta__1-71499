VERSION 5.00
Begin VB.Form frmadduser 
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2355
      FrameColor      =   16744576
      Caption         =   "Enter the password:"
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
      ColorFrom       =   16744576
      ColorTo         =   16744576
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   480
         Width           =   5175
      End
      Begin EnrollmentSystem.ChameleonBtn cmdcancel 
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Felix Titling"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15724527
         BCOLO           =   12648384
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777152
         MPTR            =   1
         MICON           =   "frmadduser.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Set rs = Louie("Select * from tbladmi_info where admi_password= '" & Trim(Me.txtPassword.Text) & "'", adUseClient, connect)
    If Not rs.EOF Then
    If USER = True And txtPassword.Text = rs!admi_password Then
        Unload Me
        frmadmi.Show vbModal
    ElseIf USER = False And txtPassword.Text = rs!admi_password Then
        Unload Me
            frmregistrar.Show vbModal
    Else
        MsgBox "Sorry you do not have a access!!!!", vbExclamation, "Info"
        Unload Me
    End If
    End If
End If
End Sub
