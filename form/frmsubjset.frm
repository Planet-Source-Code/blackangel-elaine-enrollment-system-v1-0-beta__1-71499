VERSION 5.00
Begin VB.Form frmsubjset 
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   5055
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   14843
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Subject Information:"
         Height          =   6495
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5055
         Begin VB.ComboBox cboyearlevel 
            Height          =   345
            ItemData        =   "frmsubjset.frx":0000
            Left            =   1800
            List            =   "frmsubjset.frx":001F
            TabIndex        =   6
            Text            =   "Click--->>>"
            Top             =   360
            Width           =   3015
         End
         Begin VB.ComboBox cbosubjname 
            Height          =   345
            ItemData        =   "frmsubjset.frx":007C
            Left            =   1920
            List            =   "frmsubjset.frx":00A7
            TabIndex        =   5
            Text            =   "Click--->>>"
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Height          =   1335
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1560
            Width           =   2895
         End
         Begin VB.ComboBox cbosubjteacher 
            Height          =   345
            Left            =   240
            TabIndex        =   3
            Text            =   "Click--->>>"
            Top             =   3480
            Width           =   4575
         End
         Begin VB.ComboBox cbosubjtime 
            Height          =   345
            ItemData        =   "frmsubjset.frx":0119
            Left            =   1920
            List            =   "frmsubjset.frx":014A
            TabIndex        =   2
            Text            =   "Click--->>>"
            Top             =   4200
            Width           =   2895
         End
         Begin VB.ComboBox cbotime 
            Height          =   345
            ItemData        =   "frmsubjset.frx":01CE
            Left            =   1920
            List            =   "frmsubjset.frx":01EA
            TabIndex        =   1
            Text            =   "Click--->>>"
            Top             =   4800
            Width           =   2895
         End
         Begin EnrollmentSystem.ChameleonBtn ChameleonBtn2 
            Height          =   495
            Left            =   2760
            TabIndex        =   7
            Top             =   5640
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Felix Titling"
               Size            =   9.75
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
            MICON           =   "frmsubjset.frx":026A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin EnrollmentSystem.ChameleonBtn ChameleonBtn1 
            Height          =   495
            Left            =   1680
            TabIndex        =   8
            Top             =   5640
            Width           =   1005
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&add"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Felix Titling"
               Size            =   9.75
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
            MICON           =   "frmsubjset.frx":0286
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin EnrollmentSystem.ChameleonBtn ChameleonBtn3 
            Height          =   495
            Left            =   3960
            TabIndex        =   15
            Top             =   5640
            Width           =   1005
            _ExtentX        =   3201
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&exit"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Felix Titling"
               Size            =   9.75
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
            MICON           =   "frmsubjset.frx":02A2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Name:"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Description:"
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Day:"
            Height          =   375
            Left            =   -120
            TabIndex        =   12
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Year Level:"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Teacher:"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            Height          =   375
            Left            =   1200
            TabIndex        =   9
            Top             =   4800
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmsubjset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbosubjname_Click()
Me.Text1.Text = vbNullString
Me.Text1.SetFocus
Set rs = Louie("Select * from tblf_info where faculty_department = '" & Trim(Me.cbosubjname) & "'  ", adUseClient, connect)
If Not rs.EOF Then
    rs.MoveFirst
        Do While Not rs.EOF
            Me.cbosubjteacher.AddItem rs!faculty_Surname
            rs.MoveNext
        Loop
        End If
End Sub






Private Sub ChameleonBtn1_Click()
If Me.cboyearlevel.Text = "Click--->>>" Then
    MsgBox "Year Level?", vbCritical, "Info"
    Exit Sub
End If
If Me.cbosubjname.Text = "Click--->>>" Then
    MsgBox "Subject?", vbCritical, "Info"
    Exit Sub
End If
If Me.Text1.Text = vbNullString Then
    MsgBox "Subject Description?", vbCritical, "Info"
    Exit Sub
End If
If Me.cbosubjteacher.Text = "Click--->>>" Then
    MsgBox "Teacher Subject", vbCritical, "Info"
    Exit Sub
End If
If Me.cbosubjtime.Text = "Click--->>>" Then
    MsgBox "Subject Day?", vbCritical, "Info"
    Exit Sub
End If
If Me.cbotime.Text = "Click--->>>" Then
    MsgBox "Subject Time?", vbCritical, "Info"
    Exit Sub
End If
Set rs = Louie("select* from tblshcd where subject ='" & Me.cboyearlevel.Text & "-" & Me.cbosubjname.Text & "'", adUseClient, connect)
        If rs.EOF Then
            rs.AddNew
            rs!yrlevel = Me.cboyearlevel.Text
            rs!subject = Me.cboyearlevel.Text & "-" & Me.cbosubjname.Text
            rs!Description = Me.Text1.Text
            rs!teacher = Me.cbosubjteacher.Text
            rs!Day = Me.cbosubjtime.Text
            rs!Time = Me.cbotime.Text
            rs.Update
        MsgBox "Subject Added!", vbInformation, ""
        frmsubject.kjk
        empty_obj
        Else
        MsgBox " Subject Already have a Schedule", vbInformation, ""
        empty_obj
        End If
        
        

End Sub
Public Sub empty_obj()
With Me
    .cboyearlevel.Text = "Click--->>>"
    .cbosubjname.Text = "Click--->>>"
    .Text1.Text = vbNullString
    .cbosubjteacher.Text = "Click--->>>"
    .cbosubjtime.Text = "Click--->>>"
    .cbotime.Text = "Click--->>>"
    .cbosubjteacher.Clear
End With
End Sub

Private Sub ChameleonBtn2_Click()
empty_obj
End Sub

Private Sub ChameleonBtn3_Click()
Unload Me
End Sub
