VERSION 5.00
Begin VB.Form bout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About::."
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sag"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   4450
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   -120
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "And hey you are fully free to use it without or with credit for me.It's meant for any application::>"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "santosgurung07@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"bout.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "bout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
