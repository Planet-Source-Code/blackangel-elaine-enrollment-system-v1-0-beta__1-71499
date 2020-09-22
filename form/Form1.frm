VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txta_name 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox txtconfirmpass 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3240
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txta_password 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3240
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txta_num 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin EnrollmentSystem.ChameleonBtn cmdok 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "Form1.frx":0000
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
      Left            =   5280
      TabIndex        =   4
      Top             =   2520
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
      MICON           =   "Form1.frx":001C
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
      Left            =   7320
      TabIndex        =   5
      Top             =   2520
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
      MICON           =   "Form1.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete Name:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "AdministorNumber:"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub empty_obj()
With Me
    .txtstudentnum.Text = vbNullString
    .txtFirstName.Text = vbNullString
    .txtAge.Text = vbNullString
    .Text7.Text = "Click--->>>"
    .cmbGender.Text = "Click--->>>"
    .txtCitizenship.Text = "Filipino"
    .DTPicker1.Value = Format(Now, "mm/dd/yyyy")
    .cboschyrlevel.Text = "Click--->>>"
    .cboyrlevel.Text = "Click--->>>"
    .txtPlaceOfBirth.Text = vbNullString
    .txtHomeAddress.Text = vbNullString
    .Text2.Text = vbNullString
    .Text1.Text = vbNullString
    .Text3.Text = vbNullString
    .Text4.Text = vbNullString
    .Text6.Text = vbNullString
    .Text5.Text = vbNullString
End With
End Sub
Public Sub locked()
With Me
    .cmdok.Enabled = False
    .txtAge.Enabled = False
    .Text7.Enabled = False
    .cmbGender.Enabled = False
    .txtCitizenship.Enabled = False
    .DTPicker1.Enabled = False
    .cboschyrlevel.Enabled = False
    .cboyrlevel.Enabled = False
    .txtPlaceOfBirth.Enabled = False
    .txtHomeAddress.Enabled = False
    .Text2.Enabled = False
    .Text1.Enabled = False
    .Text3.Enabled = False
    .Text4.Enabled = False
    .Text6.Enabled = False
    .Text5.Enabled = False
End With
End Sub
Public Sub unlocked()
With Me
    .cmdok.Enabled = True
    .txtAge.Enabled = True
    .Text7.Enabled = True
    .cmbGender.Enabled = True
    .txtCitizenship.Enabled = True
    .DTPicker1.Enabled = True
    .cboschyrlevel.Enabled = True
    .cboyrlevel.Enabled = True
    .txtPlaceOfBirth.Enabled = True
    .txtHomeAddress.Enabled = True
    .Text2.Enabled = True
    .Text1.Enabled = True
    .Text3.Enabled = True
    .Text4.Enabled = True
    .Text6.Enabled = True
    .Text5.Enabled = True
End With
End Sub

Private Sub cmdok_Click()
Call edit_student
End Sub

Private Sub edit_student()
Set rs = Louie("select * from tbladmi_info where admin_number= '" & Trim(Me.txta_num.Text) & "'", adUseClient, connect)
If Not rs.EOF Then
        
        rs!admi_number = Me.txta_num.Text
        rs("admi_name") = txta_name.Text
        rs!admi_username = admi_username
        rs!admi_password = txta_password.Text
    rs.Update
    MsgBox "Student Record Sucessfully Updated!!!", vbInformation, "Info"
   empty_obj
locked
    Me.txta_num.Text.SetFocus
  Else
  MsgBox "Admin Information Already Exist!!!", vbCritical, "InfO"
   Call cmdcancel_Click

End If
End Sub

Private Sub Form_Load()
 
End Sub

Private Sub txta_num_Change()

End Sub
