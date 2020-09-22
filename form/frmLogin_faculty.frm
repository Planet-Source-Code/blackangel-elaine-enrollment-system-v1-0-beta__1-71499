VERSION 5.00
Begin VB.Form frmLogin_faculty 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1657.287
   ScaleMode       =   0  'User
   ScaleWidth      =   4830.876
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      FrameColor      =   65280
      Caption         =   "Faculty Log-In"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   12648384
      ColorTo         =   16777152
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00000000&
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         MouseIcon       =   "frmLogin_faculty.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2160
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2445
      End
      Begin VB.ComboBox cbousername 
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmLogin_faculty.frx":0152
         Left            =   1920
         List            =   "frmLogin_faculty.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00000000&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         MouseIcon       =   "frmLogin_faculty.frx":0156
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2160
         Width           =   1140
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00000000&
         Caption         =   "exit"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         MouseIcon       =   "frmLogin_faculty.frx":02A8
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   615
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1005
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Plese choose your username and enter your password in the space provided bellow to login."
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4560
         Picture         =   "frmLogin_faculty.frx":03FA
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Attempt:"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblctr 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   1680
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLogin_faculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As String
Public LoginSucceeded As Boolean
Private Sub cbousername_Click()
Me.txtPassword.SetFocus
End Sub
Private Sub cmdcancel_Click()
  Unload Me
  frmchoose_login.Show vbModal
 End Sub

Private Sub cmdexit_Click()
Exit_User = MsgBox(" Do you want to close the Enrollment System?", vbYesNo, "Close")
     If Exit_User = vbYes Then
     MsgBox "Enrollment System", vbInformation, ""
      End
     Else
   Set rs = Nothing
     Me.Show
     End If
End Sub

Private Sub cmdok_Click()

 Set rs = Louie("Select * from tblf_info where faculty_username='" & Trim(cbousername.Text) & "'", adUseClient, connect)
    If cbousername = "" And txtPassword = "" Then
        MsgBox "Complete First!!", vbCritical, ""
        Me.cbousername.SetFocus
    ElseIf cbousername = rs!faculty_username And txtPassword = rs!faculty_password Then
        LoginSucceeded = True
        MDIForm1.faculty
        With current_info
            .username = Me.cbousername.Text
            .password = Me.txtPassword.Text
            MDIForm1.StatusBar1.Panels(3).Text = .username
        End With
        Unload Me
    ElseIf cbousername = rs!faculty_username And txtPassword <> rs!faculty_password Then
        MsgBox "Invalid Password, try again!", , "Login"
         LoginSucceeded = False
         Me.txtPassword = ""
         txtPassword.SetFocus
             SendKeys "{Home}+{End}"
         ctr = ctr - 1
          Me.lblctr.Caption = ctr
         If ctr = 0 Then
         MsgBox "Sorry You already used all attempt!!!", vbExclamation, "info"
           End
         End If
End If

End Sub
Private Sub Form_Load()

Set rs = Louie("Select * from tblf_info order by faculty_username", adUseClient, connect)
        If rs.EOF Then Exit Sub
            rs.MoveFirst
                Do While Not rs.EOF
              cbousername.AddItem (rs!faculty_username)
        rs.MoveNext
    Loop
    rs.Close
ctr = 3
Me.lblctr.Caption = ctr
End Sub


