VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form delusr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete User: :."
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "delusr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4680
      Begin VB.Image Image3 
         Height          =   765
         Left            =   240
         Picture         =   "delusr.frx":169B2
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please select your username and enter your password in the space provided bellow."
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label warntxt 
         BackStyle       =   0  'Transparent
         Caption         =   $"delusr.frx":170E4
         Height          =   615
         Left            =   960
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4680
      TabIndex        =   5
      Top             =   2100
      Width           =   4680
      Begin VB.CommandButton delCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Deluser 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   5000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "delusr.frx":17170
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4680
   End
   Begin MSForms.CheckBox delCheckbx 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
      VariousPropertyBits=   746588179
      BackColor       =   16777215
      ForeColor       =   4210752
      DisplayStyle    =   4
      Size            =   "6800;661"
      Value           =   "0"
      Caption         =   "Are you sure you want to delete selected user::."
      PicturePosition =   393216
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combouser 
      Height          =   350
      Left            =   360
      TabIndex        =   1
      Top             =   1320
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
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1080
   End
End
Attribute VB_Name = "delusr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
 Call displayuser
End Sub
Private Sub delCancel_Click()
Unload Me
End Sub
Private Sub Deluser_Click()
On Error GoTo err
 If delCheckbx.Value = True And combouser.Value <> "" _
 & "" And combouser.Value <> "Select user" Then
 Lgn.MoveFirst
  While Not Lgn.EOF
   If LCase$(Lgn!UNAME) = LCase$(combouser.SelText) Then
    Lgn.Delete
    Lgn.Update
    MsgBox "User: " & combouser.SelText & " is deleted." _
    & vbCr & "Selected user records are deleted.", vbInformation, "Users"
    Unload Me
    Call Mycon.loginstate
    Exit Sub
   End If
   Lgn.MoveNext
   Wend
 Else
 MsgBox "Please check the checkfield", vbInformation
End If
Exit Sub
err:
 MsgBox "Deletion process can't forward.", vbCritical, "Problem::."
End Sub
Public Sub displayuser()
i = 0
Lgn.MoveFirst
 While Not Lgn.EOF
  If Lgn!ATYPE <> 110 Then
   combouser.AddItem Lgn!UNAME, i
   i = i + 1
  End If
  Lgn.MoveNext
 Wend
 If i <= 0 Then
  warntxt.Visible = True
  Label1.Visible = False
  combouser.Visible = False
  delusr.Height = 1995
  Deluser.Enabled = False
  delCancel.Caption = "&Ok"
 End If
End Sub
