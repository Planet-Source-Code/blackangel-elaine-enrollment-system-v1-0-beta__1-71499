VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmedit_studinfo 
   Caption         =   "Student_info"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmedit_studinfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   10140
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   16748
      FrameColor      =   65280
      Caption         =   "Student_list"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Begin VB.ComboBox cboselection 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmedit_studinfo.frx":000C
         Left            =   1200
         List            =   "frmedit_studinfo.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtentry 
         Height          =   390
         Left            =   6360
         TabIndex        =   0
         Top             =   720
         Width           =   3495
      End
      Begin EnrollmentSystem.ChameleonBtn cmdview 
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   8400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&View All"
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
         MICON           =   "frmedit_studinfo.frx":0065
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
         Left            =   8040
         TabIndex        =   5
         Top             =   8400
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
         MICON           =   "frmedit_studinfo.frx":0081
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6855
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   16
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
            Name            =   "Arial"
            Size            =   8.25
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
      Begin EnrollmentSystem.ChameleonBtn cmdprint 
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   8400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Print"
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
         MICON           =   "frmedit_studinfo.frx":009D
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
         TabIndex        =   8
         Top             =   9120
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         Height          =   7095
         Left            =   120
         Top             =   1200
         Width           =   9855
      End
      Begin VB.Label lblEnterkeyword 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblselectcritia 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Critia:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmedit_studinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboselection_Click()
Me.txtentry.Text = vbNullString
Me.txtentry.SetFocus
End Sub
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()

    If Me.cboselection.Text = "All _Student" Then
     Set rs = Louie("Select * from student_info", adUseClient, connect)
            Set DataReport1.DataSource = rs
            DataReport1.Sections(2).Controls("label3").Caption = Me.cboselection.Text
             DataReport1.Sections(2).Controls("label4").Caption = Me.txtentry.Text
            DataReport1.Sections("Section3").Controls("Label2").Caption = MDIForm1.StatusBar1.Panels(3).Text
         DataReport1.Sections("Section3").Controls("Label5").Caption = Format(Now, "mm/dd/yy")
         
        DataReport1.Show vbModal
    End If
       If Me.cboselection.Text = "Student_number" Then
    Set rs = Louie("select * from student_info where Student_number = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
           Set DataReport1.DataSource = rs
            DataReport1.Sections(2).Controls("label3").Caption = Me.cboselection.Text
             DataReport1.Sections(2).Controls("label4").Caption = Me.txtentry.Text
            DataReport1.Sections("Section3").Controls("Label2").Caption = MDIForm1.StatusBar1.Panels(3).Text
         DataReport1.Sections("Section3").Controls("Label5").Caption = Format(Now, "mm/dd/yy")
        DataReport1.Show vbModal
        End If
        If Me.cboselection.Text = "Student_Schoolyear" Then
    Set rs = Louie("select * from student_info where Student_Schoolyear = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
            Set DataReport1.DataSource = rs
            DataReport1.Sections(2).Controls("label3").Caption = Me.cboselection.Text
             DataReport1.Sections(2).Controls("label4").Caption = Me.txtentry.Text
            DataReport1.Sections("Section3").Controls("Label2").Caption = MDIForm1.StatusBar1.Panels(3).Text
         DataReport1.Sections("Section3").Controls("Label5").Caption = Format(Now, "mm/dd/yy")
        DataReport1.Show vbModal
        End If
        If Me.cboselection.Text = "Student_Yearlevel" Then
    Set rs = Louie("select * from student_info where Student_Yrlevel = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
           Set DataReport1.DataSource = rs
            DataReport1.Sections(2).Controls("label3").Caption = Me.cboselection.Text
             DataReport1.Sections(2).Controls("label4").Caption = Me.txtentry.Text
            DataReport1.Sections("Section3").Controls("Label2").Caption = MDIForm1.StatusBar1.Panels(3).Text
         DataReport1.Sections("Section3").Controls("Label5").Caption = Format(Now, "mm/dd/yy")
        DataReport1.Show vbModal
        End If
End Sub

Private Sub cmdview_Click()

Set rs = Louie("Select * From student_info", adUseClient, connect)
Set Me.DataGrid1.DataSource = rs
Me.cboselection.Text = "All _Student"
End Sub
Private Sub Form_Load()
DataGrid1.Visible = True
End Sub
Private Sub txtentry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.cboselection.Text = "Student_number" Then
    Set rs = Louie("select * from student_info where Student_number = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
            Set Me.DataGrid1.DataSource = rs
        End If
        If Me.cboselection.Text = "Student_Schoolyear" Then
    Set rs = Louie("select * from student_info where Student_Schoolyear = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
            Set Me.DataGrid1.DataSource = rs
        End If
        If Me.cboselection.Text = "Student_Yearlevel" Then
    Set rs = Louie("select * from student_info where Student_Yrlevel = '" & Trim(Me.txtentry.Text) & "'", adUseClient, connect)
            Set Me.DataGrid1.DataSource = rs
        End If
        
        End If
    End Sub
