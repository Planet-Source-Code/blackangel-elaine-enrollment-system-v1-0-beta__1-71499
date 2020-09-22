VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmschd 
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10035
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
   ScaleHeight     =   8040
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   16748
      FrameColor      =   65280
      Caption         =   "Schedule"
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
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmschd.frx":0000
         Left            =   6480
         List            =   "frmschd.frx":0002
         TabIndex        =   7
         Text            =   "Click--->>>"
         Top             =   720
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   4455
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   7858
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            BorderStyle     =   0
            ForeColor       =   8388608
            HeadLines       =   1
            RowHeight       =   18
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Felix Titling"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox cboyr_lv 
         Height          =   345
         ItemData        =   "frmschd.frx":0004
         Left            =   2520
         List            =   "frmschd.frx":0023
         TabIndex        =   1
         Text            =   "Click--->>>"
         Top             =   720
         Width           =   2535
      End
      Begin EnrollmentSystem.ChameleonBtn cmdprint 
         Height          =   495
         Left            =   6240
         TabIndex        =   4
         Top             =   6840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&print"
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
         MICON           =   "frmschd.frx":0080
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
         Left            =   8160
         TabIndex        =   5
         Top             =   6840
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
         MICON           =   "frmschd.frx":009C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   7680
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   661
         Style           =   1
         SimpleText      =   "Enrollment System"
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   1
               Enabled         =   0   'False
               Object.Width           =   20029
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Felix Titling"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Faculty Name"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   720
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9960
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Yearlevel:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmschd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cboyr_lv_Click()
Set rs = Louie("Select * from tblshcd where yrlevel = '" & Trim(Me.cboyr_lv.Text) & "'", adUseClient, connect)
Set Me.DataGrid2.DataSource = rs

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If Me.cboyr_lv.Visible = True Then
Set rs = Louie("Select * from tblshcd where yrlevel = '" & Trim(Me.cboyr_lv.Text) & "'", adUseClient, connect)
Set DataReport7.DataSource = rs
    DataReport7.Sections(2).Controls("L1").Caption = Me.cboyr_lv.Text
    DataReport7.Sections("Section3").Controls("L2").Caption = MDIForm1.StatusBar1.Panels(3).Text
      DataReport7.Sections("section3").Controls("L3").Caption = Format(Now, "mm/dd/yyyy")
      DataReport7.Show vbModal
Else
Set rs = Louie("Select * from tblshcd where teacher = '" & Trim(Me.Combo1.Text) & "'", adUseClient, connect)
    Set DataReport6.DataSource = rs
    DataReport6.Sections(2).Controls("L1").Caption = Me.Combo1.Text
    DataReport6.Sections("Section3").Controls("L2").Caption = MDIForm1.StatusBar1.Panels(3).Text
      DataReport6.Sections("section3").Controls("L3").Caption = Format(Now, "mm/dd/yyyy")
      DataReport6.Show vbModal
End If


End Sub

Private Sub Combo1_Click()
Set rs = Louie("Select * from tblshcd where teacher = '" & Trim(Me.Combo1.Text) & "'", adUseClient, connect)
Set Me.DataGrid2.DataSource = rs

End Sub

Private Sub Form_Load()
Set rs = Louie("Select * from tblshcd ", adUseClient, connect)
If Not rs.EOF Then
    rs.MoveFirst
        Do While Not rs.EOF
            Me.Combo1.AddItem rs!teacher
            rs.MoveNext
            Loop
            End If
            
End Sub

