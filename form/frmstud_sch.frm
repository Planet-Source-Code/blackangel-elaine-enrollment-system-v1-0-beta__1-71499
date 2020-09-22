VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmstud_sch 
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   16748
      FrameColor      =   6974058
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Student_Schedule"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   16777152
      ColorTo         =   12648384
      Begin VB.TextBox txtstudent_name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1200
         Width           =   7065
      End
      Begin VB.TextBox txtstudentnum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3225
      End
      Begin VB.TextBox txtyrlevel 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
      End
      Begin EnrollmentSystem.ChameleonBtn cmdprint 
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   8040
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
         MICON           =   "frmstud_sch.frx":0000
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
         Left            =   5520
         TabIndex        =   5
         Top             =   8040
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
         MICON           =   "frmstud_sch.frx":001C
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
         Left            =   7440
         TabIndex        =   6
         Top             =   8040
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
         MICON           =   "frmstud_sch.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4815
         Left            =   600
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8493
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
      Begin EnrollmentSystem.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   6480
         TabIndex        =   0
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&ok"
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
         MICON           =   "frmstud_sch.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape1 
         Height          =   5535
         Left            =   240
         Top             =   2160
         Width           =   9735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10080
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level:"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Student Complete Name:"
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmstud_sch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChameleonBtn1_Click()
Set rs = Louie("Select * from student_info where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
 If Not rs.EOF Then
        Me.txtstudent_name.Text = rs!Student_name
        Me.txtyrlevel.Text = rs!Student_yrlevel
        Call sched
        End If
        
End Sub

Private Sub cmdcancel_Click()
With Me
        .txtstudent_name.Text = vbNullString
        .txtstudentnum.Text = vbNullString
        .txtyrlevel.Text = vbNullString
        .DataGrid2.Visible = False
End With
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub sched()
Set rs = Louie("Select * from tblshcd where yrlevel ='" & Trim(Me.txtyrlevel.Text) & "'", adUseClient, connect)
Set Me.DataGrid2.DataSource = rs
dbg
End Sub
Public Sub dbg()
Me.DataGrid2.Visible = True
Me.DataGrid2.Columns(0).Visible = False
End Sub

Private Sub cmdprint_Click()
Set rs = Louie("Select * from tblshcd where yrlevel ='" & Trim(Me.txtyrlevel.Text) & "'", adUseClient, connect)
Set DataReport5.DataSource = rs
    DataReport5.Sections(2).Controls("Ld").Caption = Now
    DataReport5.Sections(2).Controls("L1").Caption = txtstudentnum.Text
    DataReport5.Sections(2).Controls("L2").Caption = txtstudent_name.Text
    DataReport5.Sections(2).Controls("L3").Caption = txtyrlevel.Text
    DataReport5.Sections("section3").Controls("L4").Caption = MDIForm1.StatusBar1.Panels(3).Text
    DataReport5.Sections("Section3").Controls("L5").Caption = Format(Now, "mm/dd/yyyy")
    DataReport5.Show vbModal
    
    
End Sub

Private Sub Form_Activate()
Me.txtstudentnum.SetFocus
End Sub


Private Sub txtstudentnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ChameleonBtn1_Click
End If
End Sub
