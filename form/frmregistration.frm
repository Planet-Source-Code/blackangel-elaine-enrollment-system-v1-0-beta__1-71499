VERSION 5.00
Begin VB.Form frmregistration 
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9255
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
   ScaleHeight     =   6375
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin enrollmentsystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   16748
      FrameColor      =   65280
      Caption         =   "FAculty_registration"
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
      Begin VB.TextBox txtconfirmpass 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox txtf_password 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox txtf_contactno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2040
         TabIndex        =   16
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtf_age 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   840
         TabIndex        =   15
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtf_address 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1560
         TabIndex        =   14
         Top             =   2520
         Width           =   7575
      End
      Begin VB.TextBox txtf_surname 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1680
         TabIndex        =   8
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtf_name 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtf_num 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin enrollmentsystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   2
         TX              =   "&OK"
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
         BCOLO           =   15724527
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmregistration.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin enrollmentsystem.ChameleonBtn cmdcancel 
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   2
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
         BCOLO           =   15724527
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmregistration.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Left            =   5160
         TabIndex        =   18
         Top             =   4200
         Width           =   3375
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
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty Number:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty Name:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   6600
         Picture         =   "frmregistration.frx":0038
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SurName:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   3600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmregistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As Recordset
Dim faculty_count, faculty_username As String

Public Sub empty_obj()
With frmregistration
    .txtf_num.Text = vbNullString
    .txtf_name.Text = vbNullString
    .txtf_surname.Text = vbNullString
    .txtf_address.Text = vbNullString
    .txtf_age.Text = vbNullString
    .txtf_contactno.Text = vbNullString
    .txtf_password.Text = vbNullString
    .txtconfirmpass.Text = vbNullString
    .cboposition.Text = "click >>>"
    
End With
End Sub
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If txtf_name.Text = vbNullString Then
MsgBox "What is the Faculty Name?", vbInformation, "Faculty Name"
txtf_name.SetFocus
Exit Sub
End If
If txtf_surname.Text = vbNullString Then
MsgBox "What is the Faculty Surname?", vbInformation, "Faculty Surname"
Me.txtf_surname.SetFocus
Exit Sub
End If
If Len(txtf_password.Text) < 6 Then
MsgBox "Maximum of 6 Character", vbCritical, ""
txtf_password.Text = vbNullString
txtf_password.SetFocus
Exit Sub
End If
If txtf_password.Text <> txtconfirmpass.Text Then
MsgBox "Confirm your Password First!!!", vbCritical, ""
txtconfirmpass.Text = vbNullString
txtconfirmpass.SetFocus
Exit Sub
End If
Set rs = CN("select * from tblf_info where faculty_number= '" & Trim(Me.txtf_num.Text) & "'", adUseClient, conn)
If rs.EOF Then
rs.AddNew
rs!faculty_number = txtf_num.Text
rs!faculty_name = txtf_name.Text
rs!faculty_surname = txtf_surname.Text
rs!faculty_address = txtf_surname.Text
rs!faculty_age = txtf_age.Text
rs!faculty_contactno = Me.txtf_contactno.Text
rs!Position = Me.cboposition.Text
faculty_username = "f" & txtf_surname.Text
rs!faculty_username = faculty_username
rs!faculty_password = txtf_password.Text
rs.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Username = " & faculty_username & Chr(13) & "Password = " & txtf_password.Text, vbInformation, "Warning")
empty_obj
Call Form_Load
Else
MsgBox " Faculty Already Exist"
End If
End Sub
Private Sub Form_Load()
Set rs = CN("select * from tblf_info order by faculty_number", adUseClient, conn)
faculty_count = rs.RecordCount + 1
txtf_num.Text = "faculty" & faculty_count
End Sub
