VERSION 5.00
Begin VB.Form frmscyr 
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      FrameColor      =   65280
      Caption         =   "Add School Yr"
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Height          =   855
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   4215
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   2400
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   495
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "to"
            BeginProperty Font 
               Name            =   "Felix Titling"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   3
            Top             =   360
            Width           =   495
         End
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   2085
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
         MICON           =   "frmscyr.frx":0000
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
End
Attribute VB_Name = "frmscyr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdok_Click()

If Len(Me.Text1.Text) <> 4 Then Exit Sub

Set rs = Louie("Select * from tblschoolyr where sy= '" & Trim(Me.Text1.Text & "-" & Me.Text2.Text) & "'", adUseClient, connect)
If rs.EOF Then
        rs.AddNew
        rs!sy = Me.Text1.Text & "-" & Me.Text2.Text
        rs.Update
          UsernameAndPasswordLastShow = MsgBox("Sucessfully Added!" & Chr(10) & Chr(10) & Chr(10) & "School Year= " & Me.Text1.Text & Chr(10) & "-" & Me.Text2.Text & Chr(10), vbInformation, "Warning")
            Me.Text1.Text = vbNullString
          Me.Text2.Text = vbNullString
          Unload Me
          frmsy.Show vbModal
          Else
          MsgBox "School Year Already Exist!!!", vbCritical, "Info"
          Unload Me
          End If
         
End Sub

Private Sub Text1_Change()
Dim a As String
a = (Val(Me.Text1.Text)) + (Val(1))
Me.Text2.Text = a
End Sub
