VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmsubject 
   ClientHeight    =   8610
   ClientLeft      =   -405
   ClientTop       =   165
   ClientWidth     =   6210
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
   ScaleHeight     =   8610
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   8655
      Left            =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15266
      FrameColor      =   65280
      Caption         =   "List of Subject"
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
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   7200
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   7080
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Filter Records"
         Height          =   735
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4815
         Begin VB.ComboBox cboyr_lv 
            Height          =   345
            ItemData        =   "frmsubject.frx":0000
            Left            =   2640
            List            =   "frmsubject.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Year Level:"
            BeginProperty Font 
               Name            =   "Felix Titling"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2895
         End
      End
      Begin EnrollmentSystem.ChameleonBtn cmdexit 
         Height          =   495
         Left            =   5040
         TabIndex        =   3
         Top             =   7680
         Width           =   1005
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Exit"
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
         MICON           =   "frmsubject.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmddelete 
         Height          =   495
         Left            =   3960
         TabIndex        =   4
         Top             =   7680
         Width           =   1005
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Delete"
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
         MICON           =   "frmsubject.frx":0098
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
         Height          =   6015
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12648384
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
            Name            =   "Felix Titling"
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
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   8280
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
      Begin EnrollmentSystem.ChameleonBtn ChameleonBtn1 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   7680
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Subject Sched Setting"
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
         MICON           =   "frmsubject.frx":00B4
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
         Height          =   6255
         Left            =   120
         Top             =   1320
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmsubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboyr_lv_Click()

kjk
End Sub
Public Sub kjk()
On Error GoTo errorhandler

Set rs = Louie("Select  * from tblshcd where yrlevel ='" & Trim(Me.cboyr_lv.Text) & "'", adUseClient, connect)
Set Me.DataGrid1.DataSource = rs
db
Exit Sub
errorhandler:
MsgBox "No Schedule Subject Exist!!!", vbExclamation, ""
End Sub
Public Sub db()
Me.DataGrid1.Columns(0).Visible = False
End Sub

Private Sub ChameleonBtn1_Click()
frmsubjset.Show vbModal
End Sub

Private Sub cmddelete_Click()
On Error Resume Next
a = MsgBox("Do you want to delete?", vbYesNo, "Confirm")
If a = vbYes Then
Set rs = Louie("Select * from tblshcd where subject ='" & Trim(Me.Text2.Text) & "'", adUseClient, connect)
rs.Delete

MsgBox "Data is deleted", vbInformation
Call cboyr_lv_Click

End If

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
Set rs = Louie("Select * from tblshcd where yrlevel = '" & Trim(Me.cboyr_lv.Text) & "'", adUseClient, connect)
Text1.Text = rs!Id

Text2.Text = rs!subject



End Sub

Private Sub Form_Load()
kjk
End Sub

