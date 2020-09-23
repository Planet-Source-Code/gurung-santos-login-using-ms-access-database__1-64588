VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login: :::."
   ClientHeight    =   2520
   ClientLeft      =   150
   ClientTop       =   375
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer fadeform 
      Interval        =   25
      Left            =   120
      Top             =   2040
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.Image Image3 
         Height          =   765
         Left            =   240
         Picture         =   "Login.frx":0CCA
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select your username and enter your password in the space provided bellow."
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   3195
      End
   End
   Begin VB.CommandButton LgnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton LgnOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin MSForms.TextBox lgnuser 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Enter your username::."
      Top             =   1080
      Width           =   2925
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "5159;609"
      Value           =   "sag"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox lgnpass 
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Enter password::."
      Top             =   1440
      Width           =   2925
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "5159;609"
      PasswordChar    =   42
      Value           =   "123"
      SpecialEffect   =   3
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "Login.frx":13FC
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1515
      Width           =   960
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4570
      Y1              =   1935
      Y2              =   1935
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jk As Integer
Private Sub Form_Load()
 Login.WindowState = 0
 Call Mycon.loginstate
 jk = 0
  Transparent.ofFrm Login.hWnd, 0
End Sub
Private Sub LgnCancel_Click()
 jk = MsgBox("Are you sure you want to exit." _
 & "", vbQuestion + vbYesNo, "Confirm..")
 If jk = vbNo Then
 Exit Sub
 End If
 Unload Me
 End
End Sub
Private Sub LgnOk_Click()
 Call Mycon.loginstate
 lgnuser.Text = LCase$(Trim(lgnuser.Text))
 lgnpass.Text = LCase$(Trim(lgnpass.Text))
 If lgnuser.Text = "" Or lgnpass.Text = "" Then
 MsgBox "Blank Username or Password is restricted ", vbInformation
   Exit Sub
 End If
 While Not Lgn.EOF
  If (LCase$(lgnuser.Text) = LCase$(Lgn!UNAME)) And _
  (LCase$(lgnpass.Text) = LCase$(Lgn!PASSWORD)) Then
       Rps.StatusBar.Panels(3) = lgnuser.Text
       If Lgn!ATYPE = 110 Then
        Rps.StatusBar.Panels(4) = "Administrator"
        Rps.mnucreateuser.Enabled = True
        Rps.mnuDeluser.Enabled = True
        Else
        Rps.StatusBar.Panels(4) = "Limited user"
        Rps.mnucreateuser.Enabled = False
        Rps.mnuDeluser.Enabled = False
       End If
      GoTo cRRECT
  End If
 Lgn.MoveNext
 Wend
    MsgBox ("Invalid username or password")
 lgnuser.SetFocus
 lgnpass.Text = ""
 Exit Sub
cRRECT:
 Unload Me
      'Enable the main mdi form
 Rps.Enabled = True
End Sub

Private Sub fadeform_Timer()
If jk <= 100 Then
 Transparent.ofFrm Login.hWnd, jk
jk = jk + 20
Else
fadeform.Enabled = False
Transparent.ofFrm Login.hWnd, 255
End If
End Sub
