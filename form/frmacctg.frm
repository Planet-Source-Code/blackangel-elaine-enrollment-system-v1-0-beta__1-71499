VERSION 5.00
Begin VB.Form frmacctg 
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11295
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
   ScaleHeight     =   5115
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   16748
      FrameColor      =   16744576
      Caption         =   "Accountng_user"
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
      Begin VB.TextBox txtac_num 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtac_name 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         TabIndex        =   4
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox txac_address 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         TabIndex        =   3
         Top             =   2040
         Width           =   7455
      End
      Begin VB.TextBox txtac_contactno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         TabIndex        =   2
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtac_password 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtconfirmpass 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   3840
         Width           =   3255
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&add"
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
         MICON           =   "frmacctg.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmdcancel 
         Height          =   495
         Left            =   7200
         TabIndex        =   7
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
         MICON           =   "frmacctg.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmdexit 
         Height          =   495
         Left            =   9240
         TabIndex        =   8
         Top             =   4440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Exit"
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
         MICON           =   "frmacctg.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1440
         TabIndex        =   16
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   8400
         Picture         =   "frmacctg.frx":0054
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2160
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Name:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Accting User Number:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "plz. complete the information......"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "(Maximim 6 character)"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Top             =   3360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmacctg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim user_count, user_username As String

Public Sub empty_obj()
With frmacctg
    .txtac_num.Text = vbNullString
    .txtac_name.Text = vbNullString
    .txac_address.Text = vbNullString
    .txtac_contactno.Text = vbNullString
    .txtac_password.Text = vbNullString
    .txtconfirmpass.Text = vbNullString
End With
End Sub
Private Sub cmdcancel_Click()
empty_obj
Call Form_Load
End Sub

Private Sub cmdok_Click()
If txtac_name.Text = vbNullString Then
    MsgBox "What is the User_ Complete Name?", vbInformation, "User_Name"
    txtac_name.SetFocus
    Exit Sub
End If
If Len(txtac_password.Text) < 6 Then
    MsgBox "Maximum of 6 Character", vbCritical, "Password"
    txtac_password.Text = vbNullString
    txtac_password.SetFocus
    Exit Sub
End If
If txtac_password.Text <> txtconfirmpass.Text Then
    MsgBox "Confirm your Password First!!!", vbCritical, ""
    txtconfirmpass.Text = vbNullString
    txtconfirmpass.SetFocus
    Exit Sub
End If
Set rs = Louie("Select * from tblacct_info where Accting_number= '" & Trim(Me.txtac_num.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
        rs!Accting_number = txtac_num.Text
        rs!Accting_name = txtac_name.Text
        rs!Accting_address = txac_address.Text
        rs!Accting_contactno = Me.txtac_contactno.Text
        user_username = "acct-" & txtac_name.Text
        rs!accting_username = user_username
        rs!Accting_password = txtac_password.Text
    rs.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Username = " & user_username & Chr(13) & "Password = " & txtac_password.Text, vbInformation, "Warning")
empty_obj
Call Form_Load
Else
MsgBox " User_Already Exist"
End If
End Sub
Private Sub cmdexit_Click()
Dim a As String
a = MsgBox("Do you to exit in this Form?", vbYesNo, "Info")
If a = vbYes Then
Unload Me
Else
End If
End Sub

Private Sub Form_Load()
Set rs = Louie("select * from tblacct_info order by Accting_number", adUseClient, connect)
user_count = rs.RecordCount + 1
txtac_num.Text = "actnuser" & user_count
End Sub

