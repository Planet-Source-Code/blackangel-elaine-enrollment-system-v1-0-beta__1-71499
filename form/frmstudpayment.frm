VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmstudpayment 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10065
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
   ScaleHeight     =   8490
   ScaleWidth      =   10065
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   16748
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
      Begin VB.TextBox txtyrlevel 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtstudentnum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   3225
      End
      Begin VB.TextBox txtstudent_name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1200
         Width           =   7065
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Width           =   9855
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   2055
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3625
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
      Begin VB.TextBox txtcharges 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtpayment 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         TabIndex        =   3
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtbalance 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3480
         Width           =   2295
      End
      Begin EnrollmentSystem.ChameleonBtn cmdprint 
         Height          =   495
         Left            =   4320
         TabIndex        =   7
         Top             =   5280
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
         MICON           =   "frmstudpayment.frx":0000
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
         Left            =   6240
         TabIndex        =   8
         Top             =   5280
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
         MICON           =   "frmstudpayment.frx":001C
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
         TabIndex        =   9
         Top             =   5280
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
         MICON           =   "frmstudpayment.frx":0038
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
         Left            =   3240
         TabIndex        =   10
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&compute"
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
         MICON           =   "frmstudpayment.frx":0054
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
         Height          =   375
         Left            =   6240
         TabIndex        =   18
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
         MICON           =   "frmstudpayment.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Complete Name:"
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level:"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9960
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total charges:"
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment:"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance:"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   3600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmstudpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChameleonBtn1_Click()
If Len(Me.txtstudentnum.Text) <= 0 Then Exit Sub
Set rs = Louie("Select * from tblstudent_account where student_number ='" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
If Not rs.EOF Then
    Me.txtstudent_name.Text = rs!Student_name
    Me.txtyrlevel.Text = rs!Student_yrlevel
    Me.txtcharges = rs!temporary
    Call sss
    Else
    MsgBox "No Student Info Exist!!!", vbCritical, "Info"
    empty_obj
    End If
End Sub

Private Sub cmdcancel_Click()
empty_obj
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub


Private Sub cmdok_Click()
Dim a, s, f As String

a = Me.txtpayment.Text
f = Me.txtcharges.Text
If txtpayment.Text = vbNullString Then
   txtpayment.Text = 0

End If
Me.txtbalance.Text = Val(f) - Val(a)
Call update_info
End Sub

Private Sub cmdprint_Click()
Set rs = Louie("Select * from tblstudent_account", adUseClient, connect)
Set DataReport4.DataSource = rs
With DataReport4
    .Sections("Section2").Controls("lbldate").Caption = Now
    .Sections("Section2").Controls("L1").Caption = Me.txtstudentnum.Text
    .Sections("Section2").Controls("L2").Caption = Me.txtstudent_name.Text
    .Sections("section2").Controls("L3").Caption = Me.txtyrlevel.Text
    .Sections("Section2").Controls("L4").Caption = Me.txtcharges.Text
    .Sections("Section2").Controls("L5").Caption = Me.txtpayment.Text
    .Sections("section2").Controls("L6").Caption = Me.txtbalance.Text
    .Sections("Section2").Controls("label12").Caption = MDIForm1.StatusBar1.Panels(3).Text
    .Sections("Section2").Controls("label13").Caption = Format(Now, "mm/dd/yyyy")
    .Show vbModal
End With
End Sub

Private Sub Form_Activate()
Me.txtstudentnum.SetFocus
End Sub


Private Sub txtstudentnum_KeyPress(KeyAscii As Integer)
If Len(Me.txtstudentnum.Text) <= 0 Then Exit Sub
If KeyAscii = 13 Then
Call ChameleonBtn1_Click
End If
End Sub
Public Sub empty_obj()
With Me
Me.DataGrid2.Visible = False
.txtbalance.Text = vbNullString
.txtcharges.Text = vbNullString
.txtpayment.Text = vbNullString
.txtstudent_name.Text = vbNullString
.txtstudentnum.Text = vbNullString
.txtyrlevel.Text = vbNullString
.txtstudentnum.SetFocus
End With
End Sub
Private Sub update_info()
Dim a As String
Set rs = Louie("Select * from tblstudent_account where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
    If Not rs.EOF Then
        rs!temporary = txtbalance.Text
        rs.Update
    If Not rs.EOF Then
        rs.AddNew
        rs!Student_Number = Me.txtstudentnum.Text
        rs!Student_name = Me.txtstudent_name.Text
        rs!Student_yrlevel = Me.txtyrlevel.Text
        rs!Student_balance = Me.txtbalance.Text
        rs!student_dateofpayment = Now
        rs!Student_currentpayment = Me.txtpayment.Text
        rs.Update
        
        a = MsgBox("Automatically Updated!!!!,Do  you want to Print?", vbYesNo, "info")
        If a = vbYes Then
         Call sss
        Call cmdprint_Click
        Else
          Call sss
        Call cmdcancel_Click
        End If
        Else
        empty_obj
        End If
     End If
End Sub
Public Sub dbg()
Me.DataGrid2.Visible = True
Me.DataGrid2.Columns(0).Visible = False
Me.DataGrid2.Columns(1).Visible = False
Me.DataGrid2.Columns(2).Visible = False
Me.DataGrid2.Columns(4).Visible = False
End Sub
Private Sub sss()
Set rs = Louie("Select * from tblstudent_account where student_number = '" & Trim(Me.txtstudentnum.Text) & "'", adUseClient, connect)
Set Me.DataGrid2.DataSource = rs
dbg
End Sub
