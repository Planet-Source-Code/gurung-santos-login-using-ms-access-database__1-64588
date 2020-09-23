VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Rps 
   BackColor       =   &H8000000C&
   Caption         =   "Result processing system ::."
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10740
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6990
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "Rps.frx":0000
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7329
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "Rps.frx":039C
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1/24/2004"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:16 PM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuadm 
      Caption         =   "&Administrator"
      Begin VB.Menu mnucreateuser 
         Caption         =   "&Create user"
      End
      Begin VB.Menu mnuDeluser 
         Caption         =   "&Delete user"
      End
      Begin VB.Menu mnublank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChngpassword 
         Caption         =   "C&hange password"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "&Log off user"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "Rps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnucreateuser_Click()
dInt = 1
Adduser.Show vbModal, Me
End Sub
Private Sub mnuDeluser_Click()
delusr.Show vbModal, Me
End Sub
Private Sub mnuChngpassword_Click()
cngPass.Show vbModal, Me
End Sub
Private Sub mnulogoff_Click()
Login.Show vbModal, Me
End Sub

Private Sub MDIForm_Load()
Me.Show
Me.Enabled = False
   Lgn.Open "Select * from login", con, adOpenDynamic, adLockOptimistic
   If Lgn.BOF = True Then
      'Lgn.Close
      dInt = 110 'FOR THE FIRST TIME , SO ADMINISTRATOR CREATED
      Adduser.Show vbModal, Me
   Else
      'Lgn.Close
      dInt = 1
      Login.Show vbModal, Me
   End If
End Sub
