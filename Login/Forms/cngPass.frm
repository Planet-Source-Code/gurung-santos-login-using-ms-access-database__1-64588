VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form cngPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change password::."
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "cngPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4935
      TabIndex        =   7
      Top             =   3795
      Width           =   4935
      Begin VB.CommandButton delCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Deluser 
         Caption         =   "&Change"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
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
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5040
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "cngPass.frx":0CCA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   765
         Left            =   360
         Picture         =   "cngPass.frx":2AA24
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"cngPass.frx":2B156
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   3795
      End
   End
   Begin MSForms.TextBox cngNpass 
      Height          =   345
      Left            =   600
      TabIndex        =   12
      Top             =   2640
      Width           =   3735
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6588;609"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password::."
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
      Index           =   3
      Left            =   600
      TabIndex        =   11
      Top             =   3000
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "New password::."
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
      TabIndex        =   10
      Top             =   2400
      Width           =   2040
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password::."
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1800
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1080
   End
   Begin MSForms.TextBox cngCnpass 
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6588;609"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox cngCpass 
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "6588;609"
      PasswordChar    =   42
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combouser 
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   7
      Size            =   "6800;617"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "Select user"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "cngPass.frx":2B1EF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   5040
   End
End
Attribute VB_Name = "cngPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cc As Integer
Private Sub Deluser_Click()
cngCpass.Text = LCase$(Trim(cngCpass.Text))
cngCpass.Text = LCase$(Trim(cngCpass.Text))
cngCpass.Text = LCase$(Trim(cngCpass.Text))
If cngNpass.Text <> cngCnpass.Text Then
 MsgBox "The password you typed donot match." _
 & "Please retype the new password in both boxes", vbExclamation, "user account"
 cngNpass.Text = ""
 cngCnpass.Text = ""
 Exit Sub
End If
If cngCpass = "" Or cngNpass = "" Or cngCnpass = "" Then
 MsgBox "A blank password is strictly restricted." _
 & "Please type the passwords correctly:.", vbExclamation, "user account"
 cngCpass.Text = ""
 cngNpass.Text = ""
 cngCnpass.Text = ""
Exit Sub
End If
Lgn.MoveFirst
While Not Lgn.EOF
 cc = Lgn!ATYPE
 If combouser.SelText = LCase$(Lgn!UNAME) And cngCpass.Text = LCase$(Lgn!PASSWORD) Then
   'Lgn.Delete
   'Lgn.AddNew
   Lgn!UNAME = combouser.SelText
   Lgn!PASSWORD = cngNpass.Text
   If cc = 110 Then
    Lgn!ATYPE = 110 'Administrator no. is 110
   Else
    Lgn!ATYPE = 1   'Limited user no. is 1
   End If
   Lgn.Update
MsgBox "The password of " & combouser.SelText & " has been changed" _
& ".", vbInformation, "User account"
   Unload Me
   Exit Sub
 End If
 Lgn.MoveNext
Wend
If combouser.SelText = "Select user" Then
MsgBox "Please select a username from the combo box" _
& ".", vbExclamation, "User account"
Else
MsgBox "Password of " & combouser.SelText & " is incorrect" _
& ".", vbExclamation, "User account"
cngCpass.Text = ""
cngCnpass.Text = ""
cngNpass.Text = ""
End If
End Sub
Private Sub Form_Load()
 Call displayuser
End Sub
Private Sub delCancel_Click()
Unload Me
End Sub
Public Sub displayuser()
Lgn.MoveFirst
 While Not Lgn.EOF
   combouser.AddItem LCase$(Lgn!UNAME)
   Lgn.MoveNext
 Wend
End Sub
