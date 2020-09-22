VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmaddpar 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6615
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Felix Titling"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Compute"
      Height          =   375
      Left            =   4560
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
   End
   Begin EnrollmentSystem.jcFrames jcFrames1 
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11668
      FrameColor      =   16744576
      BackColor       =   16777215
      Caption         =   "Add Particular"
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
      ColorFrom       =   16777215
      ColorTo         =   16777215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Particular Information"
         Height          =   4815
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   6135
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2880
            TabIndex        =   15
            Top             =   3240
            Width           =   2895
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2880
            TabIndex        =   14
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2880
            TabIndex        =   13
            Top             =   2280
            Width           =   2895
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2880
            TabIndex        =   12
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2880
            TabIndex        =   11
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtamount 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2280
            TabIndex        =   6
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox txtparticular 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   2040
            TabIndex        =   5
            Top             =   840
            Width           =   2895
         End
         Begin VB.ComboBox cboyr_lv 
            Height          =   390
            ItemData        =   "frmaddpar.frx":0000
            Left            =   1800
            List            =   "frmaddpar.frx":001F
            TabIndex        =   4
            Text            =   "Click--->>>>"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "computer fee (sti)"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Tuition Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Books"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Miscellaneous fee"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Registration Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Year Level:"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Particular:"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
            Height          =   255
            Left            =   1080
            TabIndex        =   1
            Top             =   3840
            Width           =   1575
         End
      End
      Begin EnrollmentSystem.ChameleonBtn cmdok 
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   5520
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
         MICON           =   "frmaddpar.frx":007C
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
         Left            =   4320
         TabIndex        =   8
         Top             =   5520
         Width           =   1005
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Cancel"
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
         MICON           =   "frmaddpar.frx":0098
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
         Left            =   5400
         TabIndex        =   9
         Top             =   5520
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
         MICON           =   "frmaddpar.frx":00B4
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
         TabIndex        =   10
         Top             =   6240
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
   End
End
Attribute VB_Name = "frmaddpar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub enabled_obj()
txtparticular.Enabled = True
txtamount.Enabled = True
End Sub
Public Sub disabled()
txtparticular.Enabled = False
txtamount.Enabled = False
End Sub
Private Sub cboyr_lv_Click()
enabled_obj
Me.txtparticular.SetFocus
End Sub
Private Sub cmdcancel_Click()
disabled
Me.cboyr_lv.Text = "Click--->>>"
Me.txtparticular.Text = vbNullString
Me.txtamount.Text = vbNullString
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdok_Click()
If Me.txtparticular.Text = vbNullString Then
    MsgBox "Particular Item?", vbCritical, "Info"
    txtparticular.SetFocus
    Exit Sub
End If
If txtamount.Text = vbNullString Then
    MsgBox "Amount of Particular item?", vbCritical, "Info"
    txtamount.SetFocus
    Exit Sub
End If
Set rs = Louie("Select * from tblparticular where particular = '" & Trim(cboyr_lv.Text & "-" & txtparticular.Text) & "'", adUseClient, connect)
If rs.EOF Then
    rs.AddNew
    rs!yearlevel = cboyr_lv.Text
    rs!particular = cboyr_lv.Text & "-" & txtparticular.Text
    rs!Registration_Fee = Text1.Text
    rs!Miscellaneous_Fee = Text2.Text
    rs!Book_Fee = Text3.Text
    rs!Tuition_Fee = Text4.Text
    rs!Computer_Fee = Text5.Text
    rs!amount = txtamount.Text
    rs.Update
     
     UsernameAndPasswordLastShow = MsgBox("Sucessfully Added!" & Chr(10) & Chr(10) & Chr(10) & "Particular Item:= " & txtparticular.Text & Chr(10) & " Item Amount:=" & txtamount.Text, vbInformation, "Warning")
     
     disabled
     
Me.cboyr_lv.Text = "Click--->>>"
Me.txtparticular.Text = vbNullString
Me.txtamount.Text = vbNullString

    Else
    
    MsgBox "Particular item Already Exist!", vbCritical, "Info"
    disabled
Me.cboyr_lv.Text = "Click--->>>"
Me.txtparticular.Text = vbNullString
Me.txtamount.Text = vbNullString
     End If
End Sub

Private Sub Command1_Click()
txtamount.Text = Val(Text1) + Val(Text2) + Val(Text3) + Val(Text4) + Val(Text5)
End Sub

Private Sub Form_Load()
disabled
End Sub
