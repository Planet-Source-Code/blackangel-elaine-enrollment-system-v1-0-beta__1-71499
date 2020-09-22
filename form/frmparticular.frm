VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmparticular 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6270
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   6510
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   6840
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2280
      Width           =   1455
   End
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   11033
      FrameColor      =   65280
      Caption         =   "Particulars"
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
      Begin VB.TextBox txttotalpayment 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         TabIndex        =   6
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Filter Records"
         Height          =   855
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   6135
         Begin VB.ComboBox cboyr_lv 
            Height          =   405
            ItemData        =   "frmparticular.frx":0000
            Left            =   2760
            List            =   "frmparticular.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   2895
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
            TabIndex        =   1
            Top             =   360
            Width           =   2895
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4683
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
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   5280
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
         MICON           =   "frmparticular.frx":007C
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
         Left            =   5160
         TabIndex        =   8
         Top             =   5280
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
         MICON           =   "frmparticular.frx":0098
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
         Left            =   4080
         TabIndex        =   9
         Top             =   5280
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
         MICON           =   "frmparticular.frx":00B4
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
         TabIndex        =   11
         Top             =   5880
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
      Begin VB.Shape Shape1 
         Height          =   3135
         Left            =   120
         Top             =   1560
         Width           =   6255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment:"
         Height          =   615
         Left            =   1440
         TabIndex        =   5
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Total Payment:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmparticular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String


Private Sub cboyr_lv_Click()
On Error GoTo errorhandler
Set rs = Louie("Select * from tblparticular where yearlevel='" & Trim(Me.cboyr_lv.Text) & "' order by particular asc", adUseClient, connect)
Set Me.DataGrid1.DataSource = rs
 dbgrid
On Error GoTo errorhandler
Set rs = Louie("Select sum(amount) as sumpay from tblparticular where yearlevel ='" & Trim(Me.cboyr_lv.Text) & "'", adUseClient, connect)
Me.txttotalpayment.Text = rs!sumpay
Set rs = Nothing
Exit Sub
errorhandler:
Me.txttotalpayment = 0
End Sub

Private Sub cmddelete_Click()

a = MsgBox("Do you want to delete?", vbYesNo, "Confirm")
If a = vbYes Then
Set rs = Louie("Select * from tblparticular where particular ='" & Trim(Text2.Text) & "'", adUseClient, connect)
rs.Delete
MsgBox "Data is deleted", vbInformation
Set rs = Nothing
Call cboyr_lv_Click
End If

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
frmaddpar.Show vbModal
End Sub

Public Sub dbgrid()
DataGrid1.Columns(0).Visible = False
DataGrid1.Columns(1).Width = 1500
DataGrid1.Columns(2).Width = 2300
DataGrid1.Columns(3).Width = 1000
End Sub

Private Sub DataGrid1_Click()

On Error Resume Next
Set rs = Louie("Select * from tblparticular where yearlevel='" & Trim(Me.cboyr_lv.Text) & "' order by particular asc", adUseClient, connect)

Text2.Text = rs!particular

End Sub



Private Sub Form_Load()
DataGrid1.Visible = True
End Sub
