VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   957.15
   ScaleMode       =   0  'User
   ScaleWidth      =   4647.782
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcFrames jcFrames1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      Caption         =   "Login"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2445
      End
      Begin VB.ComboBox cbousername 
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmLogin1.frx":0000
         Left            =   1920
         List            =   "frmLogin1.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3600
         Top             =   2520
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\THESIS\database\mbms.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\THESIS\database\mbms.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   615
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1005
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Set rs = DbConnect("Select * from login order by username", adUseClient, conn)
If rs.EOF Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
    cbousername.AddItem (rs!UserName)
    rs.MoveNext
Loop
rs.Close
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set rs = DbConnect("Select * from login where username='" & cbousername & "'", adUseClient, conn)
If cbousername = "" And txtPassword = "" Then
MsgBox "Complete First!!", vbCritical, ""
Me.cbousername.SetFocus
ElseIf cbousername = rs!UserName And txtPassword = rs!Password Then
        LoginSucceeded = True
        Unload Me
       frmregistration.Show
ElseIf cbousername = rs!UserName And txtPassword <> rs!Password Then
        MsgBox "Sorry!You do not have a acess to register!!!", , "Login"
         LoginSucceeded = False
         Unload Me
         MDIForm1.Enabled = True
End If
End If
End Sub
Private Sub cbousername_Click()
Me.txtPassword.SetFocus
End Sub
