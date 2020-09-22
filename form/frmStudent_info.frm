VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Student_info 
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   1065
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   15240
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   2400
      TabIndex        =   42
      Top             =   2640
      Width           =   2775
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
      Height          =   405
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   40
      Top             =   6600
      Width           =   3225
   End
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9375
      Left            =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16536
      FrameColor      =   6974058
      BackColor       =   16777215
      FillColor       =   16777215
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Student Registration Form"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   16777215
      ColorTo         =   16777215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "If Transferee"
         Height          =   1575
         Left            =   240
         TabIndex        =   24
         Top             =   7200
         Width           =   12255
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   11
            Top             =   360
            Width           =   7425
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   12
            Top             =   840
            Width           =   7425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous School:"
            Height          =   285
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous School Address:"
            Height          =   285
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   3000
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   6855
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   12255
         Begin VB.ComboBox cboschyrlevel 
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
            Left            =   6720
            TabIndex        =   7
            Text            =   "Click--->>>"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   3975
            Left            =   120
            TabIndex        =   25
            Top             =   2760
            Width           =   12015
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Document Presented"
               Height          =   2175
               Left            =   9240
               TabIndex        =   46
               Top             =   240
               Width           =   2655
               Begin VB.CheckBox Check2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Form 137"
                  Height          =   495
                  Left            =   240
                  TabIndex        =   49
                  Top             =   960
                  Width           =   2175
               End
               Begin VB.CheckBox Check1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Birth Certificate"
                  Height          =   495
                  Left            =   240
                  TabIndex        =   48
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.CheckBox Check3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Form 138"
                  Height          =   495
                  Left            =   240
                  TabIndex        =   47
                  Top             =   1560
                  Width           =   2175
               End
            End
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
               Height          =   405
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   38
               Top             =   3000
               Width           =   4305
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
               Height          =   405
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   36
               Top             =   2040
               Width           =   4305
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
               Height          =   405
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   35
               Top             =   2520
               Width           =   4305
            End
            Begin VB.TextBox txtHomeAddress 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1245
               Left            =   2160
               MaxLength       =   50
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   720
               Width           =   4305
            End
            Begin VB.TextBox txtPlaceOfBirth 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   9
               Top             =   240
               Width           =   6585
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   3480
               Width           =   1935
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Guardian's Name"
               Height          =   285
               Left            =   120
               TabIndex        =   37
               Top             =   3000
               Width           =   1965
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mother's Name"
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   2640
               Width           =   1710
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Father's Name "
               Height          =   285
               Left            =   120
               TabIndex        =   33
               Top             =   2160
               Width           =   1710
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Home Address"
               Height          =   435
               Left            =   120
               TabIndex        =   27
               Top             =   720
               Width           =   1995
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Place Of Birth"
               Height          =   435
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   2415
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6720
            TabIndex        =   6
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            Format          =   56295425
            CurrentDate     =   39666
         End
         Begin VB.ComboBox cboyrlevel 
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
            ItemData        =   "frmStudent_info.frx":0000
            Left            =   6720
            List            =   "frmStudent_info.frx":001F
            TabIndex        =   8
            Text            =   "Click--->>>"
            Top             =   2280
            Width           =   2655
         End
         Begin VB.ComboBox cmbGender 
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
            ItemData        =   "frmStudent_info.frx":007C
            Left            =   6750
            List            =   "frmStudent_info.frx":0086
            TabIndex        =   4
            Text            =   "Click--->>>"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtCitizenship 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "Filipino"
            Top             =   840
            Width           =   2625
         End
         Begin VB.TextBox txtAge 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1800
            Width           =   1065
         End
         Begin VB.TextBox txtLastName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1320
            Width           =   2745
         End
         Begin VB.TextBox txtMiddleName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   1
            Top             =   840
            Width           =   2745
         End
         Begin VB.TextBox txtFirstName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   0
            Top             =   360
            Width           =   2745
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School Yr:"
            Height          =   285
            Left            =   4920
            TabIndex        =   30
            Top             =   1800
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "YRS OLD"
            Height          =   285
            Left            =   3360
            TabIndex        =   23
            Top             =   1800
            Width           =   1290
         End
         Begin VB.Label Religion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Religion"
            Height          =   435
            Left            =   720
            TabIndex        =   22
            Top             =   2280
            Width           =   1515
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   435
            Left            =   5040
            TabIndex        =   21
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            Height          =   435
            Left            =   5040
            TabIndex        =   20
            Top             =   840
            Width           =   1725
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yr. Level"
            Height          =   285
            Left            =   5280
            TabIndex        =   19
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   315
            Left            =   5400
            TabIndex        =   18
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   1800
            Width           =   765
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   315
            Left            =   480
            TabIndex        =   16
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   2460
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   555
            Left            =   480
            TabIndex        =   14
            Top             =   360
            Width           =   2205
         End
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   9120
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   661
         Style           =   1
         SimpleText      =   "WORD Enrollment System"
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   1
               Enabled         =   0   'False
               Object.Width           =   26802
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtsnum 
         Height          =   405
         Left            =   480
         TabIndex        =   31
         Top             =   9600
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   12720
         TabIndex        =   41
         Top             =   480
         Width           =   2295
         Begin EnrollmentSystem.ChameleonBtn cmdcancel 
            Height          =   495
            Left            =   240
            TabIndex        =   43
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            MICON           =   "frmStudent_info.frx":0098
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin EnrollmentSystem.ChameleonBtn cmdok 
            Height          =   495
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&add"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            MICON           =   "frmStudent_info.frx":00B4
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
            TabIndex        =   45
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Exit"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            MICON           =   "frmStudent_info.frx":00D0
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
End
Attribute VB_Name = "Student_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comfirm, student_count, a, ex As String
Private Sub count_stud()
Set rs = Louie("Select * from student_info order by Student_number", adUseClient, connect)
        student_count = rs.RecordCount + 1
        Me.txtsnum.Text = Format(Now(), "yy-") & "E00" & student_count
End Sub
Public Sub empty_obj()
With Me
    .txtFirstName.Text = vbNullString
    .txtMiddleName.Text = vbNullStringb
    .txtLastName.Text = vbNullString
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

Private Sub add_student()
  Call count_stud
Set rs = Louie("select * from student_info where Student_number= '" & Trim(Me.txtsnum.Text) & "'", adUseClient, connect)

If rs.EOF Then
    rs.AddNew
        rs!Student_Number = Me.txtsnum.Text
        rs!Student_name = Me.txtLastName.Text & "," & Me.txtFirstName.Text & " " & Me.txtMiddleName.Text
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
    UsernameAndPasswordLastShow = MsgBox("Student Number and Student Name!" & Chr(13) & Chr(13) & Chr(13) & "Student Number = " & Me.txtsnum.Text & Chr(13) & "Student Name  = " & Me.txtLastName.Text & "," & Me.txtFirstName.Text & " " & Me.txtMiddleName.Text, vbInformation, "Warning")
   empty_obj
  Else
  MsgBox "Student Information Already Exist!!!", vbCritical, "InfO"
    empty_obj
    
End If


        
        
End Sub

Private Sub cmdcancel_Click()
empty_obj
count_stud

End Sub

Private Sub cmdexit_Click()
Unload Me

End Sub

Private Sub cmdok_Click()
If Me.txtFirstName.Text = vbNullString Then
    MsgBox "Student FirstName?", vbCritical, "Info"
    txtFirstName.SetFocus
Exit Sub
End If
If Me.txtMiddleName.Text = vbNullString Then
    MsgBox "Student MiddleName?", vbCritical, "Info"
    txtMiddleName.SetFocus
Exit Sub
End If
If Me.txtLastName.Text = vbNullString Then
    MsgBox "Student LastName?", vbCritical, "Info"
    txtLastName.SetFocus
Exit Sub
End If
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
If Me.Text1.Text = vbNullString And Me.Text2.Text = vbNullString Then
confirm = MsgBox("the Student Tranferee?", vbYesNo, "confirmation")
If confrim = vbYes Then
Me.Text2.SetFocus
Exit Sub
ElseIf confirm = vbNo Then
Call add_student
Exit Sub
End If
Else
Call add_student
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
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text3_Change()

End Sub
