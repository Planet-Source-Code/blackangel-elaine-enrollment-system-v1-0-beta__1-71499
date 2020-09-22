VERSION 5.00
Begin VB.Form frmregistrar 
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10695
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
   ScaleHeight     =   5565
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   16748
      FrameColor      =   65280
      Caption         =   "Registrar_user"
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
      Begin VB.TextBox txtconfirmpass 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   4200
         Width           =   3255
      End
      Begin VB.TextBox txtr_password 
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtr_contactno 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         TabIndex        =   4
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox txtr_age 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         TabIndex        =   3
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtr_address 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         TabIndex        =   2
         Top             =   2040
         Width           =   6735
      End
      Begin VB.TextBox txtr_name 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         TabIndex        =   1
         Top             =   1560
         Width           =   6135
      End
      Begin VB.TextBox txtr_num 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   2415
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   4800
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
         MICON           =   "frmregistrar.frx":0000
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
         TabIndex        =   8
         Top             =   4800
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
         MICON           =   "frmregistrar.frx":001C
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
         TabIndex        =   9
         Top             =   4800
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
         MICON           =   "frmregistrar.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Left            =   7200
         TabIndex        =   18
         Top             =   3720
         Width           =   3375
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
         TabIndex        =   17
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2160
         TabIndex        =   15
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar User Number:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Name:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   120
         Picture         =   "frmregistrar.frx":0054
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1920
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3120
         TabIndex        =   12
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2400
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   3120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmregistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim user_count, user_username As String

Public Sub empty_obj()
    With frmregistrar
        .txtr_num.Text = vbNullString
        .txtr_name.Text = vbNullString
        .txtr_address.Text = vbNullString
        .txtr_contactno.Text = vbNullString
        .txtr_age.Text = vbNullString
        .txtr_password.Text = vbNullString
        .txtconfirmpass.Text = vbNullString
End With
End Sub
Private Sub cmdcancel_Click()
empty_obj
Call Form_Load
End Sub

Private Sub cmdok_Click()
If txtr_name.Text = vbNullString Then
    MsgBox "What is the User_ Complete Name?", vbInformation, "User_Name"
    txtr_name.SetFocus
    Exit Sub
End If
If Len(txtr_password.Text) < 6 Then
    MsgBox "Maximum of 6 Character", vbCritical, "Password"
    txtr_password.Text = vbNullString
    txtr_password.SetFocus
    Exit Sub
End If
If txtr_password.Text <> txtconfirmpass.Text Then
    MsgBox "Confirm your Password First!!!", vbCritical, ""
    txtconfirmpass.Text = vbNullString
    txtconfirmpass.SetFocus
    Exit Sub
End If
Set rs = Louie("Select * from tblreg_info where registrar_number= '" & Trim(txtr_num.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
        rs!registrar_number = Me.txtr_num.Text
        rs!registrar_name = txtr_name.Text
        rs!registrar_address = txtr_address.Text
        rs!registrar_age = txtr_age.Text
        rs!registrar_contactno = Me.txtr_contactno.Text
        user_username = "reg-" & txtr_name.Text
        rs!registrar_username = user_username
        rs!registrar_password = txtr_password.Text
    rs.Update
UsernameAndPasswordLastShow = MsgBox("Remember your password and your username!" & Chr(13) & Chr(13) & Chr(13) & "Username = " & user_username & Chr(13) & "Password = " & txtr_password.Text, vbInformation, "Warning")
empty_obj
Call Form_Load
Else
MsgBox " User_Already Exist"
End If
End Sub
Private Sub cmdexit_Click()
Dim a As String
a = MsgBox("Do you to exit in this Form?", vbYesNo, "Info")
If a = vbYes Then
Unload Me
Else
End If
End Sub

Private Sub Form_Load()
Set rs = Louie("select * from tblreg_info order by registrar_number", adUseClient, connect)
user_count = rs.RecordCount + 1
txtr_num.Text = "reguser" & user_count
End Sub
