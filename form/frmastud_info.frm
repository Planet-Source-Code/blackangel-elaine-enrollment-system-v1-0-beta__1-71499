VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmastud_info 
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10080
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
   ScaleHeight     =   9060
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   10440
      TabIndex        =   18
      Top             =   1800
      Width           =   270
   End
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9735
      Left            =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   17171
      FrameColor      =   65280
      Caption         =   "Student_list"
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Full Payment"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   2400
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Monthly Payment"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   21
         Top             =   2400
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Quarterly Payment"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   20
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txtyrlevel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Height          =   2175
         Left            =   7200
         TabIndex        =   11
         Top             =   3240
         Width           =   2535
         Begin EnrollmentSystem.ChameleonBtn cmdprint 
            Height          =   495
            Left            =   360
            TabIndex        =   12
            Top             =   360
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
            MICON           =   "frmastud_info.frx":0000
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
            Left            =   360
            TabIndex        =   13
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
            MICON           =   "frmastud_info.frx":001C
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
            Left            =   360
            TabIndex        =   14
            Top             =   1560
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
            MICON           =   "frmastud_info.frx":0038
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
      Begin VB.TextBox txttotalpayment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   2655
         Left            =   120
         TabIndex        =   7
         Top             =   6120
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   2055
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
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
      Begin VB.TextBox txtstudentname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   7065
      End
      Begin VB.TextBox txtstudentnum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3225
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   360
         TabIndex        =   0
         Top             =   3480
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
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
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   8760
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   661
         Style           =   1
         SimpleText      =   " Enrollment System"
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
      Begin EnrollmentSystem.ChameleonBtn ChameleonBtn2 
         Height          =   375
         Left            =   6360
         TabIndex        =   22
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
         MICON           =   "frmastud_info.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         Height          =   855
         Left            =   240
         Top             =   2280
         Width           =   7575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "payment:"
         Height          =   615
         Left            =   5160
         TabIndex        =   17
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         Height          =   2055
         Left            =   240
         Top             =   3360
         Width           =   6735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total charges:"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   5520
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9840
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level:"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name:"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmastud_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim payment As String

Private Sub ChameleonBtn2_Click()
Dim a As String
On Error GoTo errorhandler
Set rs = Louie("Select * from student_info where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
If Not rs.EOF Then
Me.DataGrid1.Visible = True
    txtstudentname.Text = rs!Student_name
    Me.txtyrlevel.Text = rs!Student_yrlevel
   Set rs = Louie("Select * from tblparticular where yearlevel='" & Trim(Me.txtyrlevel.Text) & "' order by particular asc", adUseClient, connect)
Set Me.DataGrid1.DataSource = rs
 dbgrid

Set rs = Louie("Select sum(Amount) as sumpay from tblparticular where yearlevel ='" & Trim(Me.txtyrlevel.Text) & "'", adUseClient, connect)
payment = rs!sumpay
Me.txttotalpayment.Text = payment
If Option1.Value = True Then
Text2.Text = Option1.Caption
Text1.Text = payment
ElseIf Option2.Value = True Then
Text2.Text = Option2.Caption
Text1.Text = payment \ 10
ElseIf Option3.Value = True Then
Text2.Text = Option3.Caption
Text1.Text = payment \ 4
End If
Call add_info
Call sss
Else
MsgBox "Student Number Not Exist!!!", vbCritical, "Info"
End If
Exit Sub
errorhandler:
Me.txttotalpayment = 0
End Sub

Private Sub cmdcancel_Click()
Me.DataGrid2.Visible = False
Me.txtstudentnum.Text = vbNullString
Me.DataGrid1.Visible = False
Me.txttotalpayment.Text = vbNullString
Me.txtstudentname.Text = vbNullString
Me.txtyrlevel.Text = vbNullString
Me.Text1.Text = vbNullString
Me.txtstudentnum.SetFocus
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
Set rs = Louie("Select * from student_info where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
    Set DataReport2.DataSource = rs
        DataReport2.Sections(2).Controls("ld").Caption = Now
        DataReport2.Sections(2).Controls("L1").Caption = Me.txtstudentnum.Text
        DataReport2.Sections(2).Controls("L2").Caption = Me.txtstudentname.Text
        DataReport2.Sections(2).Controls("L7").Caption = Me.txtyrlevel.Text
        DataReport2.Sections(2).Controls("lt").Caption = Me.txttotalpayment.Text
        DataReport2.Sections(2).Controls("label5").Caption = Me.Text2.Text
        DataReport2.Sections(2).Controls("label6").Caption = Me.Text1.Text
        DataReport2.Sections("Section3").Controls("labelu").Caption = MDIForm1.StatusBar1.Panels(3).Text
        DataReport2.Sections("Section3").Controls("ldt").Caption = Format(Now, "mm/dd/yyyy")
        Set rs = Louie("Select * from tblparticular where yearlevel='" & Trim(Me.txtyrlevel.Text) & "' order by amount DESC ", adUseClient, connect)
    Set DataReport2.DataSource = rs
        DataReport2.Show vbModal
        
End Sub

Private Sub Form_Activate()
Me.txtstudentnum.SetFocus
End Sub

Private Sub Form_Load()

Me.DataGrid2.Visible = True

Call sss

End Sub

Private Sub txtstudentnum_KeyPress(KeyAscii As Integer)

If Len(Me.txtstudentnum.Text) <= 0 Then Exit Sub
If KeyAscii = 13 Then
Call ChameleonBtn2_Click
End If



End Sub
Public Sub dbgrid()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 1500
DataGrid1.Columns(2).Width = 2300
DataGrid1.Columns(3).Width = 1000
End Sub

Private Sub add_info()
a = MsgBox("Are you sure?", vbYesNo, "info")
    If a = vbNo Then
    Exit Sub
    Else
Set rs = Louie("Select * from tblstudent_account where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
        rs!Student_Number = Me.txtstudentnum.Text
        rs!Student_name = Me.txtstudentname.Text
        rs!Student_yrlevel = Me.txtyrlevel.Text
        rs!Student_charge = payment
        rs!temporary = payment
        rs!student_dateofpayment = Now
    rs.Update
    
    a = MsgBox("Automatically Save!!!Do you want to print", vbYesNo, "Info")
    If a = vbYes Then
      Call Form_Load
    Call cmdprint_Click
    Else
      Call Form_Load
    Call cmdcancel_Click
    End If
    End If
    End If
 

End Sub
Public Sub dbg()
Me.DataGrid2.Visible = True
Me.DataGrid2.Columns(4).Visible = False
End Sub
Private Sub sss()
Set rs = Louie("Select * from tblstudent_account where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
Set Me.DataGrid2.DataSource = rs
dbg
End Sub
