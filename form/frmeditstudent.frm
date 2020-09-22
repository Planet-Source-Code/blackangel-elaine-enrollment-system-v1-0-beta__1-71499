VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form editstudent 
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14550
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
   ScaleHeight     =   10035
   ScaleWidth      =   14550
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   2160
      TabIndex        =   37
      Top             =   2760
      Width           =   2775
   End
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   10095
      Left            =   -240
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   17806
      FrameColor      =   6974058
      BackColor       =   16777215
      FillColor       =   16777215
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Student Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16777215
      ColorTo         =   16777215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "If Transferee"
         Height          =   1335
         Left            =   360
         TabIndex        =   9
         Top             =   8520
         Width           =   10695
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   11
            Top             =   720
            Width           =   5265
         End
         Begin VB.TextBox Text2 
            Height          =   405
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   10
            Top             =   240
            Width           =   5265
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address of last School:"
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   3780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of last School:"
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3165
         End
      End
      Begin VB.TextBox txtstudentnum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3225
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   7335
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   11655
         Begin VB.TextBox txtFirstName 
            Height          =   405
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            Top             =   360
            Width           =   7545
         End
         Begin VB.TextBox txtAge 
            Height          =   405
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   1065
         End
         Begin VB.TextBox txtCitizenship 
            Height          =   405
            Left            =   6900
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Filipino"
            Top             =   960
            Width           =   2625
         End
         Begin VB.ComboBox cmbGender 
            Height          =   405
            ItemData        =   "frmeditstudent.frx":0000
            Left            =   2070
            List            =   "frmeditstudent.frx":000A
            TabIndex        =   3
            Text            =   "Click--->>>"
            Top             =   2280
            Width           =   2655
         End
         Begin VB.ComboBox cboyrlevel 
            Height          =   405
            ItemData        =   "frmeditstudent.frx":001C
            Left            =   6840
            List            =   "frmeditstudent.frx":003B
            TabIndex        =   18
            Text            =   "Click--->>>"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   4215
            Left            =   360
            TabIndex        =   15
            Top             =   3000
            Width           =   11175
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   42
               Top             =   3480
               Width           =   3105
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   41
               Top             =   3000
               Width           =   4545
            End
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   30
               Top             =   2520
               Width           =   4545
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   29
               Top             =   2040
               Width           =   4545
            End
            Begin VB.TextBox txtPlaceOfBirth 
               Height          =   405
               Left            =   2760
               MaxLength       =   50
               TabIndex        =   7
               Top             =   240
               Width           =   6345
            End
            Begin VB.TextBox txtHomeAddress 
               Height          =   1245
               Left            =   2760
               MaxLength       =   50
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   720
               Width           =   4545
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Document Presented"
               Height          =   2295
               Left            =   7680
               TabIndex        =   36
               Top             =   840
               Width           =   3375
               Begin VB.CheckBox Check2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Form 137"
                  Height          =   495
                  Left            =   480
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Birth Certificate"
                  Height          =   735
                  Left            =   480
                  TabIndex        =   44
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.CheckBox Check3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Form 138"
                  Height          =   495
                  Left            =   480
                  TabIndex        =   43
                  Top             =   1680
                  Width           =   1575
               End
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Father's Name "
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   2160
               Width           =   2295
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mother's Name"
               Height          =   285
               Left            =   120
               TabIndex        =   33
               Top             =   2640
               Width           =   2175
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Guardian's Name"
               Height          =   285
               Left            =   120
               TabIndex        =   32
               Top             =   3120
               Width           =   1965
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               Height          =   285
               Left            =   120
               TabIndex        =   31
               Top             =   3600
               Width           =   1935
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Place Of Birth"
               Height          =   435
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Home Address"
               Height          =   435
               Left            =   120
               TabIndex        =   16
               Top             =   720
               Width           =   2835
            End
         End
         Begin VB.ComboBox cboschyrlevel 
            Height          =   405
            ItemData        =   "frmeditstudent.frx":0098
            Left            =   6840
            List            =   "frmeditstudent.frx":009A
            TabIndex        =   6
            Text            =   "Click--->>>"
            Top             =   1920
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6840
            TabIndex        =   5
            Top             =   1440
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            Format          =   56295425
            CurrentDate     =   39666
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Name:"
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   2160
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            Height          =   315
            Left            =   480
            TabIndex        =   26
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   315
            Left            =   600
            TabIndex        =   25
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yr. Level"
            Height          =   285
            Left            =   5400
            TabIndex        =   24
            Top             =   2400
            Width           =   1290
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            Height          =   435
            Left            =   5160
            TabIndex        =   23
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   435
            Left            =   5160
            TabIndex        =   22
            Top             =   1500
            Width           =   1905
         End
         Begin VB.Label Religion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   435
            Left            =   600
            TabIndex        =   21
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "YRS OLD"
            Height          =   285
            Left            =   2400
            TabIndex        =   20
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Yr:"
            Height          =   285
            Left            =   5040
            TabIndex        =   19
            Top             =   1920
            Width           =   1605
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   12120
         TabIndex        =   35
         Top             =   1200
         Width           =   2415
         Begin EnrollmentSystem.ChameleonBtn cmdok 
            Height          =   495
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Update"
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
            MICON           =   "frmeditstudent.frx":009C
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
            Left            =   240
            TabIndex        =   39
            Top             =   960
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
            MICON           =   "frmeditstudent.frx":00B8
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
            Left            =   240
            TabIndex        =   40
            Top             =   1680
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
            MICON           =   "frmeditstudent.frx":00D4
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   2895
      End
   End
End
Attribute VB_Name = "editstudent"
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

Private Sub Check1_Click()

End Sub

Private Sub Check3_Click()

End Sub

Private Sub cmdcancel_Click()
empty_obj
locked
 Me.txtstudentnum.SetFocus
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdok_Click()

If Me.txtAge.Text = vbNullString Then
    MsgBox "Student Name?", vbCritical, "Info"
    txtFirstName.SetFocus
Exit Sub
End If
If Me.Text7.Text = "Click--->>>" Then
    MsgBox "Student Religion?", vbCritical, "Info"
Exit Sub
End If
If Me.cmbGender.Text = "Click--->>>" Then
    MsgBox "Student Gender?", vbCritical, "Info"
Exit Sub
End If
If Me.txtCitizenship.Text = vbNullString Then
    MsgBox "Student Gender?", vbCritical, "Info"
Exit Sub
End If
If Me.cboschyrlevel.Text = "Click--->>>" Then
    MsgBox "Student School Year?", vbCritical, "Info"
Exit Sub
End If
If Me.cboyrlevel.Text = "Click--->>>" Then
    MsgBox "Student Year Level?", vbCritical, "Info"
Exit Sub
End If
If Me.txtPlaceOfBirth.Text = vbNullString Then
    MsgBox "Student Place of Birth?", vbCritical, "Info"
    txtPlaceOfBirth.SetFocus
Exit Sub
End If
If Me.txtHomeAddress.Text = vbNullString Then
    MsgBox "Student Home Address?", vbCritical, "Info"
    txtHomeAddress.SetFocus
Exit Sub
End If







End Sub
Private Sub edit_student()
Set rs = Louie("select * from student_info where Student_number= '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
If Not rs.EOF Then
        rs!Student_Number = Me.txtstudentnum.Text
        rs!Student_name = Me.txtFirstName.Text
        rs!Student_Age = Me.txtAge.Text
        rs!student_religion = Me.Text7.Text
        rs!Student_gender = Me.cmbGender.Text
        rs!student_Citizenship = Me.txtCitizenship.Text
        rs!Student_birthdate = Me.DTPicker1.Value
        rs!Student_schoolyear = Me.cboschyrlevel.Text
        rs!Student_yrlevel = Me.cboyrlevel.Text
        rs!Student_placeofbirth = Me.txtPlaceOfBirth.Text
        rs!Student_homeaddress = Me.txtHomeAddress.Text
        rs!Student_namelastschool = Me.Text2.Text
        rs!Student_addresslastschool = Me.Text1.Text
        rs!Student_Fathersname = Me.Text4.Text
        rs!Student_Mothersname = Me.Text3.Text
        rs!Student_Guardiansname = Me.Text6.Text
        rs!Student_Contact = Me.Text5.Text
        rs!Student_BC = Me.Check1
        rs!Student_Form_137 = Me.Check2
        rs!Student_Form_138 = Me.Check3
    rs.Update
    MsgBox "Student Record Sucessfully Updated!!!", vbInformation, "Info"
   empty_obj
locked
    Me.txtstudentnum.SetFocus
  Else
  MsgBox "Student Information Already Exist!!!", vbCritical, "Info"
   Call cmdcancel_Click

End If
End Sub


Private Sub Form_Load()
Set rs = Louie("Select * from tblschoolyr  ", adUseClient, connect)
If rs.EOF Then Exit Sub
rs.MoveFirst
If Not rs.EOF Then
    Do While Not rs.EOF
    Me.cboschyrlevel.AddItem rs!sy
    rs.MoveNext
    Loop
    End If
locked

End Sub


Private Sub jcFrames1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtFirstName_Change()

End Sub

Private Sub txtstudentnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(Me.txtstudentnum.Text) <= 0 Then Exit Sub
Set rs = Louie("select * from student_info where Student_number= '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
If Not rs.EOF Then
        unlocked
        Me.txtstudentnum.Text = rs!Student_Number
        Me.txtFirstName.Text = rs!Student_name
        Me.txtAge.Text = rs!Student_Age
        Me.Text7.Text = rs!student_religion
        Me.cmbGender.Text = rs!Student_gender
        Me.txtCitizenship.Text = rs!student_Citizenship
        Me.DTPicker1.Value = rs!Student_birthdate
        Me.cboschyrlevel.Text = rs!Student_schoolyear
        Me.cboyrlevel.Text = rs!Student_yrlevel
        Me.txtPlaceOfBirth.Text = rs!Student_placeofbirth
        Me.txtHomeAddress.Text = rs!Student_homeaddress
        Me.Text2.Text = rs!Student_namelastschool
        Me.Text1.Text = rs!Student_addresslastschool
        Me.Text4.Text = rs!Student_Fathersname
        Me.Text3.Text = rs!Student_Mothersname
        Me.Text6.Text = rs!Student_Guardiansname
        Me.Text5.Text = rs!Student_Contact
        Me.Check1.Caption = rs!Student_BC
        Me.Check2.Caption = rs!Student_Form_137
        Me.Check3.Caption = rs!Student_Form_138
        Else
        MsgBox "Student Information Not Exist!!!", vbCritical, "InfO"
        Call cmdcancel_Click
        End If
End If
End Sub
