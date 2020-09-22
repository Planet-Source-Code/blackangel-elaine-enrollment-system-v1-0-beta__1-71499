VERSION 5.00
Begin VB.Form frmchoose_login 
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   2250
   ClientTop       =   9210
   ClientWidth     =   10950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2566
      FrameColor      =   65280
      Caption         =   "Select:"
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
      Begin EnrollmentSystem.ChameleonBtn cmdadmin 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Administrator"
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
         MICON           =   "frmchoose_login.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmdreg 
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Registrar"
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
         MICON           =   "frmchoose_login.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmdacctng 
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "A&ccounting"
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
         MICON           =   "frmchoose_login.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EnrollmentSystem.ChameleonBtn cmdfaculty 
         Height          =   495
         Left            =   8160
         TabIndex        =   4
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Faculty"
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
         MICON           =   "frmchoose_login.frx":0054
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
   Begin EnrollmentSystem.ChameleonBtn ChameleonBtn1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Administrator"
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
      MICON           =   "frmchoose_login.frx":0070
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
Attribute VB_Name = "frmchoose_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdacctng_Click()
Unload Me
frmLogin_acct.Show vbModal
End Sub

Private Sub cmdadmin_Click()
Unload Me
frmLogin_admin.Show vbModal
End Sub

Private Sub cmdfaculty_Click()
Unload Me
frmLogin_faculty.Show vbModal
End Sub

Private Sub cmdreg_Click()
Unload Me
frmLogin_Registrar.Show vbModal
End Sub
