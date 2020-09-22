VERSION 5.00
Begin VB.Form frmfaculty 
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10335
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
   ScaleHeight     =   6795
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   16748
      FrameColor      =   65280
      Caption         =   "FAculty"
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
      Begin VB.ComboBox cbodepartment 
         Height          =   405
         ItemData        =   "frmrfacultyfrm.frx":0000
         Left            =   2760
         List            =   "frmrfacultyfrm.frx":0025
         TabIndex        =   22
         Text            =   "Click--->>>"
         Top             =   3960
         Width           =   3255
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   3120
         TabIndex        =   18
         Top             =   6000
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
         MICON           =   "frmrfacultyfrm.frx":008B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtconfirmpass 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   5400
         Width           =   3255
      End
      Begin VB.TextBox txtf_password 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   4560
         Width           =   3255
      End
      Begin VB.TextBox txtf_contactno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         TabIndex        =   14
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox txtf_age 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         TabIndex        =   13
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtf_address 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         TabIndex        =   12
         Top             =   2520
         Width           =   7455
      End
      Begin VB.TextBox txtf_surname 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         TabIndex        =   8
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtf_name 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
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
      Begin EnrollmentSystem.ChameleonBtn cmdcancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   19
         Top             =   6000
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
         MICON           =   "frmrfacultyfrm.frx":00A7
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
         Left            =   6960
         TabIndex        =   20
         Top             =   6000
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
         MICON           =   "frmrfacultyfrm.frx":00C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex. ""Mr or Mrs"" Surname"
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
         Left            =   6840
         TabIndex        =   23
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Teach:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   2295
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
         Left            =   6240
         TabIndex        =   16
         Top             =   4680
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
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   720
         TabIndex        =   10
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   4680
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
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   7680
         Picture         =   "frmrfacultyfrm.frx":00DF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SurName:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   3600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmfaculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New Recordset
Dim faculty_count, faculty_username As String

Public Sub empty_obj()
With frmfaculty
    .txtf_num.Text = vbNullString
    .txtf_name.Text = vbNullString
    .txtf_surname.Text = vbNullString
    .txtf_address.Text = vbNullString
    .txtf_age.Text = vbNullString
    .txtf_contactno.Text = vbNullString
    .txtf_password.Text = vbNullString
    .txtconfirmpass.Text = vbNullString
    .cbodepartment.Text = "click >>>"
    
End With
End Sub
Private Sub cmdcancel_Click()

empty_obj
Call Form_Load

End Sub

Private Sub cmdok_Click()
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
Set rs = Louie("select * from tblf_info where faculty_number= '" & Trim(Me.txtf_num.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
        rs!faculty_number = txtf_num.Text
        rs!faculty_name = txtf_name.Text
        rs!faculty_Surname = txtf_surname.Text
        rs!faculty_address = txtf_surname.Text
        rs!faculty_age = txtf_age.Text
        rs!faculty_contactno = Me.txtf_contactno.Text
        rs!faculty_department = Me.cbodepartment.Text
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
Private Sub cmdexit_Click()
Dim a As String
a = MsgBox("Do you to exit in this Form?", vbYesNo, "Info")
If a = vbYes Then
Unload Me
Else
End If
End Sub

Private Sub Form_Load()
Set rs = Louie("select * from tblf_info order by faculty_number", adUseClient, connect)
faculty_count = rs.RecordCount + 1
txtf_num.Text = "faculty" & faculty_count
End Sub

