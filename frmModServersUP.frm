VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmModServersUP 
   Caption         =   "Modify Server User/Password"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   Icon            =   "frmModServersUP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5318
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   2415
   End
End
Attribute VB_Name = "frmModServersUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim olddata As String
Dim ChangeInProgress As Boolean

Private Sub Command1_Click()

frmModServersUP.MousePointer = 11
Label1.Caption = "Checking Connections..."
conncheck = CheckConnections(msg$)
Label1.Caption = ""
If conncheck Then
   Set conn = New ADODB.Connection
   Set cmd1 = New ADODB.Command
   strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   Sql = "Delete from SqlDeveloperServers"
   cmd1.CommandText = Sql
   cmd1.Execute
   For z = 1 To flex.Rows - 1
      flex.Row = z
      flex.Col = 1
      sn$ = Trim(flex.Text)
      flex.Col = 2
      su$ = Trim(flex.Text)
      flex.Col = 3
      sp$ = Trim(flex.Text)
      Sql = "Insert into SqlDeveloperServers (ServerName, ServerUser, ServerPassword) values ('" + sn$ + "', '" + su$ + "', '" + sp$ + "')"
      cmd1.CommandText = Sql
      cmd1.Execute
   Next
   conn.Close
   frmModServersUP.MousePointer = 1
   Unload frmModServersUP
Else
   frmModServersUP.MousePointer = 1
   MsgBox msg$
End If

End Sub

Private Sub Command2_Click()

Unload frmModServersUP

End Sub

Private Sub Form_Load()

flex.Cols = 4
flex.ColWidth(0) = 0
flex.ColWidth(1) = 2000
flex.ColWidth(2) = 900
flex.ColWidth(3) = 2280

flex.Row = 0
flex.Col = 1
flex.Text = "Server Name/IP"
flex.Col = 2
flex.Text = "Username"
flex.Col = 3
flex.Text = "Password"

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Select * from SqlDeveloperServers order by Servername"
cmd1.CommandText = Sql
Set rs = cmd1.Execute
cr& = 0
Do While Not rs.EOF
   cr& = cr& + 1
   flex.Rows = cr& + 1
   flex.Row = cr&
   flex.Col = 1
   flex.Text = rs!ServerName
   flex.Col = 2
   flex.Text = rs!ServerUser
   flex.Col = 3
   flex.Text = rs!ServerPassword
   rs.MoveNext
Loop
conn.Close

End Sub
Private Sub flex_Click()

If flex.MouseRow = 0 Or flex.MouseCol = 1 Then Exit Sub
olddata = flex.Text
Text1.Text = olddata
Text1.SelStart = 0
Text1.Height = flex.CellHeight
Text1.Width = flex.CellWidth
Text1.Move flex.CellLeft + flex.Left, flex.CellTop + flex.Top, flex.CellWidth, flex.CellHeight
Text1.Visible = True
Text1.SetFocus
ChangeInProgress = True

End Sub

Private Sub flex_LeaveCell()

If Text1.Text = "" And ChangeInProgress Then
  flex.Text = olddata
  Text1.Visible = False
End If

If Text1.Visible Then
  flex.Text = Text1.Text
End If

Text1.Visible = False
Text1.Text = ""
ChangeInProgress = False

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
   Case 13
      flex.Text = Text1.Text
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
   Case 27
      flex.Text = olddata
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
End Select

End Sub

Public Function CheckConnections(msg$) As Boolean

On Error GoTo errorfound

CheckConnections = True
msg$ = ""
For z = 1 To flex.Rows - 1
   Tester = True
   flex.Row = z
   flex.Col = 1
   sn$ = flex.Text
   flex.Col = 2
   su$ = flex.Text
   flex.Col = 3
   sp$ = flex.Text
   Set conn = New ADODB.Connection
   Set cmd1 = New ADODB.Command
   strConnect = "Provider=SQLOLEDB;server=" + sn$ + ";uid=" + su$ + ";pwd=" + sp$ + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   conn.Close
Next
On Error GoTo 0
Exit Function

errorfound:
   CheckConnections = False
   Tester = False
   msg$ = msg$ + "Server " + sn$ + " cannot connect." + vbCrLf
   Resume Next
   
End Function
