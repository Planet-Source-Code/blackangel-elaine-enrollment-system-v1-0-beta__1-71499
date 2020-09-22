VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Security Check"
   ClientHeight    =   9285
   ClientLeft      =   2355
   ClientTop       =   705
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5485.889
   ScaleMode       =   0  'User
   ScaleWidth      =   9549.08
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   3960
      Picture         =   "frmLogins.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   10200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6960
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   5040
      Width           =   2565
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   5775
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4320
         Picture         =   "frmLogins.frx":22BB
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   8535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rights As String
Dim pass As String
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to exit now?", vbYesNo) = vbYes Then
End
End If
End Sub


Private Sub cmdOK_Click()
Dim test As String
Dim auth As Date
auth = "06 / 06 / 2007"
If auth = Date Then
MsgBox "Your Licence is Expired, Pls Contact your Administrator to give You the Licence key."
End
Else
With rs_user
 login = Combo1.Text
.MoveFirst
While Not .EOF

 'pass = .Fields(2)
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
  rights = .Fields(3)
 pass = (.Fields(1))
    test = Val(Text1.Text - 1)
    If txtPassword.Text <> pass Then
   MsgBox "Invalid Password,you have just " + test + " login trails left, input the correct password or exit", vbInformation + vbOKOnly, "Authentication"
    Text1.Text = Val(Text1.Text) - 1
  
  If Text1.Text = "0" Then
  MsgBox "Sorry, but you cant be too smart", vbCritical + vbOKOnly
 End
End If
        Exit Sub
    ElseIf txtPassword.Text = pass Then
    'If Combo2.Text = rights Then
        MsgBox "A c c e s s  G r a n t e d   " & Time, vbOKOnly, "Authentication"
        txtPassword.Text = ""
   
        Me.Hide
         If rights = 1 Then
            user1
      ElseIf rights = 2 Then
            user2
        
      ElseIf rights = 3 Then
            user3
            End If
    With rs_userlog
       .AddNew
       .Fields(0) = login
       .Fields(1) = "Log In"
       .Fields(2) = Date
       .Fields(3) = Time
       .Fields(4) = "Successful"
       .Update
    End With
       MDIMain.Show
     ' Frm_welcome.Show
        'Exit Sub
   End If
    'End If
End If
.MoveNext
Wend
End With
        End If
End Sub


Private Sub Form_Load()
Call connect
Label4.Caption = "JBTechnologies"
Label5.Caption = "Welcome to  " & Schname & ",                                " & SchAdd

With rs_user
    While Not .EOF
   Combo1.AddItem .Fields(0)
   .MoveNext
  Wend
 End With
 Text1.Text = 3
 Label3.Caption = "Only authorised users are allowed to login." & vbCrLf & "If you forgot your password, please contact the system administrator immediately."

End Sub


Private Sub user1()
        MDIMain.StatusBar1.Panels(1) = "User Name :- " & login

End Sub
Private Sub user2()
 MDIMain.StatusBar1.Panels(1) = "User Name :- " & login
       MDIMain.mnuuser.Enabled = False
        MDIMain.adm.Enabled = False
        MDIMain.mnuset.Enabled = False
End Sub

Private Sub user3()
MDIMain.StatusBar1.Panels(1) = "User Name :- " & login
       'main_menu.mnuemp.Enabled = False
       'main_menu.mnureport.Enabled = False
       'main_menu.mnu.Enabled = False
       'main_menu.mnuset.Enabled = False

End Sub

