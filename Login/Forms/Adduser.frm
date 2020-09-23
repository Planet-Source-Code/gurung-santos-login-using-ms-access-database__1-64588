VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Adduser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create user: :::."
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Adduser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4680
      Begin VB.Image Image3 
         Height          =   765
         Left            =   240
         Picture         =   "Adduser.frx":0CCA
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remember, Create user with a unique name and password so that it will be easy in future."
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   3195
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4650
      TabIndex        =   7
      Top             =   3435
      Width           =   4650
      Begin VB.CommandButton CruCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton crunext 
         Caption         =   "&Create"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   5000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5000
         Y1              =   22
         Y2              =   22
      End
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "Adduser.frx":13FC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5000
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5000
      Y1              =   1920
      Y2              =   1920
   End
   Begin MSForms.TextBox cruconfirm 
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   3615
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6376;617"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox crupass 
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6376;617"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox cruname 
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6376;617"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   1680
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
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1080
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
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   1080
   End
End
Attribute VB_Name = "Adduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.WindowState = 0
End Sub
Private Sub CruCancel_Click()
If dInt = 110 Then
End
Else
Unload Me
End If
End Sub
Private Sub crunext_Click()
cruname.Text = LCase$(Trim(cruname.Text))
crupass.Text = LCase$(Trim(crupass.Text))
cruconfirm.Text = LCase$(Trim(cruconfirm.Text))
If cruname.Text = "" Or crupass.Text = "" Or cruconfirm.Text = "" Then
   MsgBox "A blank username or password is not restricted."
   Exit Sub
End If
If crupass.Text <> cruconfirm.Text Then
   MsgBox "The Password you typed do not match." _
   & "Please retype both password correctly.", vbExclamation, "User account"
   crupass.Text = ""
   cruconfirm.Text = ""
   Exit Sub
End If
Call Mycon.loginstate
If Lgn.BOF = False Then
 Lgn.MoveFirst
 While Not Lgn.EOF
  If LCase$(Trim$(Lgn!UNAME)) = cruname.Text Then
   MsgBox "User name already exist.Please type another name." _
   & "", vbExclamation, "Username exist"
   cruname.Text = ""
   crupass.Text = ""
   cruconfirm.Text = ""
   Exit Sub
  End If
  Lgn.MoveNext
 Wend
End If
Lgn.AddNew
Lgn!UNAME = cruname.Text
Lgn!PASSWORD = crupass.Text
Lgn!ATYPE = dInt ''if dint=110 then administrator not then limited..
Lgn.Update
Call Mycon.loginstate
If dInt = 110 Then
 Rps.StatusBar.Panels(3) = cruname.Text
 Rps.StatusBar.Panels(4) = "Administrator"
 Rps.mnuadm.Enabled = True
MsgBox "Administrator created.", vbInformation, "User account"
Else
MsgBox "Limited user created.", vbInformation, "User account"
End If
Unload Me
'Lgn.Close
Rps.Enabled = True
End Sub

