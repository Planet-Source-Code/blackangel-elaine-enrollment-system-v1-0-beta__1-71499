VERSION 5.00
Begin VB.Form frmadmi 
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10890
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
   ScaleHeight     =   6330
   ScaleWidth      =   10890
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   120
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   16748
      FrameColor      =   16744576
      Caption         =   "Administrator"
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
      ColorFrom       =   16744576
      ColorTo         =   16744576
      Begin VB.TextBox txta_num 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txta_name 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         TabIndex        =   1
         Top             =   1560
         Width           =   6135
      End
      Begin VB.TextBox txta_address 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         TabIndex        =   2
         Top             =   2040
         Width           =   7575
      End
      Begin VB.TextBox txta_age 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         TabIndex        =   3
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txta_contactno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3240
         TabIndex        =   4
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox txta_password 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   4200
         Width           =   3255
      End
      Begin VB.TextBox txtconfirmpass 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   4800
         Width           =   3255
      End
      Begin VB.ComboBox cboa_position 
         Height          =   405
         ItemData        =   "frmadmi.frx":0000
         Left            =   3240
         List            =   "frmadmi.frx":000D
         TabIndex        =   5
         Text            =   "Click--->>>"
         Top             =   3600
         Width           =   3855
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   5520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&add"
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
         MICON           =   "frmadmi.frx":0035
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
         Left            =   5280
         TabIndex        =   9
         Top             =   5520
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
         MICON           =   "frmadmi.frx":0051
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
         Left            =   7320
         TabIndex        =   10
         Top             =   5520
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
         MICON           =   "frmadmi.frx":006D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2400
         TabIndex        =   18
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   8520
         Picture         =   "frmadmi.frx":0089
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Name:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   17
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "AdministorNumber:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   4920
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "plz. complete the information......"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "(Maximim 6 character)"
         BeginProperty Font 
            Name            =   "Felix Titling"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   12
         Top             =   4320
         Width           =   3375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "position:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   3600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmadmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim admi_count, admi_username As String
Private Sub cmdexit_Click()
Dim a As String
a = MsgBox("Do you to exit in this Form?", vbYesNo, "Info")
If a = vbYes Then
Unload Me
Else
End If
End Sub

Public Sub empty_obj()
With frmadmi
    .txta_address.Text = vbNullString
    .txta_contactno.Text = vbNullString
    .txta_name.Text = vbNullString
    .txta_num.Text = vbNullString
    .txta_password.Text = vbNullString
    .txta_age.Text = vbNullString
    .txtconfirmpass.Text = vbNullString
    .cboa_position.Text = "Click>>>"
End With
End Sub
Private Sub cmdcancel_Click()
empty_obj
Call Form_Load
End Sub

Private Sub cmdok_Click()
If txta_name.Text = vbNullString Then
    MsgBox "What is the Administrator Complete Name?", vbInformation, "User_Name"
    txta_name.SetFocus
    Exit Sub
End If
If cboa_position.Text = "Click--->>>" Then
    MsgBox "Your position Plzzz!!!", vbCritical, "Info"
    Exit Sub
End If

If Len(txta_password.Text) < 6 Then
    MsgBox "Maximum of 6 Character", vbCritical, "Password"
    txta_password.Text = vbNullString
    txta_password.SetFocus
    Exit Sub
End If
If txta_password.Text <> txtconfirmpass.Text Then
    MsgBox "Confirm your Password First!!!", vbCritical, ""
    txtconfirmpass.Text = vbNullString
    txtconfirmpass.SetFocus
    Exit Sub
End If
Set rs = Louie("Select * from tbladmi_info where admi_number= '" & Trim(Me.txta_num.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
        rs!admi_number = Me.txta_num.Text
        rs("admi_name") = txta_name.Text
        rs!admi_address = Me.txta_address.Text
        rs!admi_contactno = Me.txta_contactno.Text
        rs!admi_position = Me.cboa_position.Text
        admi_username = "admin-" & txta_name.Text
        rs!admi_username = admi_username
        rs!admi_password = txta_password.Text
    rs.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Username = " & admi_username & Chr(13) & "Password = " & txta_password.Text, vbInformation, "Warning")
empty_obj
Call Form_Load
Else
MsgBox " User_Already Exist"
End If
End Sub

Private Sub Form_Load()
Set rs = Louie("select * from tbladmi_info order by admi_number", adUseClient, connect)
admi_count = rs.RecordCount + 1
txta_num.Text = "admi" & admi_count
End Sub


Private Sub jcFrames1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
