VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmuserinfo 
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10200
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
   ScaleHeight     =   4665
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
      FrameColor      =   65280
      Caption         =   "User_information"
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
         Height          =   405
         ItemData        =   "frmuserinfo.frx":0000
         Left            =   360
         List            =   "frmuserinfo.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3735
      End
      Begin EnrollmentSystem.ChameleonBtn cmdcancel 
         Height          =   375
         Left            =   8040
         TabIndex        =   0
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         MICON           =   "frmuserinfo.frx":0043
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
         Height          =   2895
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5106
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
      Begin VB.Shape Shape1 
         Height          =   3135
         Left            =   240
         Top             =   960
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmuserinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
If Me.Combo1.Text = "Administrator" Then
Set rs = Louie("select * from tbladmi_info", adUseClient, connect)
Set frmuserinfo.DataGrid1.DataSource = rs
dgb1
End If
If Me.Combo1.Text = "Registrar" Then
Set rs = Louie("select * from tblreg_info", adUseClient, connect)
Set frmuserinfo.DataGrid1.DataSource = rs
dgb2
End If
If Me.Combo1.Text = "Accounting" Then
Set rs = Louie("select * from tblacct_info", adUseClient, connect)
Set frmuserinfo.DataGrid1.DataSource = rs
dgb3
End If
If Me.Combo1.Text = "Faculty" Then
Set rs = Louie("select * from tblf_info", adUseClient, connect)
Set frmuserinfo.DataGrid1.DataSource = rs
dgb4
End If
End Sub
Public Sub dgb1()
Me.DataGrid1.Columns(6).Visible = False
Me.DataGrid1.Columns(7).Visible = False
End Sub
Public Sub dgb2()
Me.DataGrid1.Columns(5).Visible = False
Me.DataGrid1.Columns(6).Visible = False
End Sub
Public Sub dgb3()
Me.DataGrid1.Columns(4).Visible = False
Me.DataGrid1.Columns(5).Visible = False
End Sub
Public Sub dgb4()
Me.DataGrid1.Columns(7).Visible = False
Me.DataGrid1.Columns(8).Visible = False
End Sub
