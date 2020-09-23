Attribute VB_Name = "Mycon"
Option Explicit
'Dim con As Connection
'Dim Lgn As ADODB.Recordset
   Public con As New ADODB.Connection
   Public Lgn As New ADODB.Recordset
   Public Gbal As Integer
   Public Gstring As String
   Public dInt As Integer
   Public guser As String
Dim Gs As String
 
Sub main()
On Error Resume Next
 If App.PrevInstance Then
   MsgBox "An instance of Rps is already running!" & vbCr _
   & "You cannot run two instances of this application at the same time", vbCritical, "Application already running"
   End
 Else
   Call getcon
   Load Rps
   Rps.Show
 End If
End Sub
Public Sub getcon()
On Error GoTo suRedb
   con.Open "Provider=Microsoft.jet.oledb.4.0;data source=" _
   & App.Path & "\Database\mydb.mdb; jet oledb:database password=123"
   Exit Sub
suRedb:
   Gs = "Either Database doesnot exist or" + vbCr
   Gs = Gs + "Database password has changed:."
   MsgBox Gs, vbCritical
   End
End Sub
Public Sub loginstate()
 If Lgn.State = 1 Then
  Lgn.Close
  Lgn.Open "Select * from login", con, adOpenDynamic, adLockOptimistic
 End If
End Sub
