VERSION 5.00
Begin VB.Form frmMoveStoredProcs 
   Caption         =   "Copy Stored Procedures"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   Icon            =   "frmMoveStoredProcs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.ComboBox cmbDB 
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Database"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "To Server"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmMoveStoredProcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currprocs(5000) As String
Dim cmax As Integer

Private Sub cmbServer_Change()

cmbServer_Click

End Sub

Private Sub cmbServer_Click()

Set conn = New adodb.Connection
Set cmd1 = New adodb.Command

strConnect = "Provider=SQLOLEDB;server=" + cmbServer.Text + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "sp_databases"
cmd1.CommandText = Sql
Set rs = cmd1.Execute

cmbDB.Clear
Do While Not rs.EOF
   cmbDB.AddItem rs!Database_Name
   rs.MoveNext
Loop
conn.Close
cmbDB.Text = cmbDB.List(0)

End Sub

Private Sub Command1_Click()

flag% = 0
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      flag% = 1
      Exit For
   End If
Next

If flag% = 0 Then
   MsgBox "No Stored Procedures have been selected."
   Exit Sub
End If

sysbut% = MsgBox("Are you sure you want to copy the selected stored procedures?", 4, "Copy Confirm")
If sysbut% <> vbYes Then Exit Sub

frmMoveStoredProcs.MousePointer = 11
Set conn = New adodb.Connection
Set cmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + cmbServer.Text + ";database=" + cmbDB.Text + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

cmd1.CommandText = "sp_help"
Set rs = cmd1.Execute
rs.Filter = "Object_type='stored procedure'"
cmax = 0
Do While Not rs.EOF
    cmax = cmax + 1
    currprocs(cmax) = rs!Name
    rs.MoveNext
Loop

numcopied = 0
numreplaced = 0
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      sp$ = List1.List(z)
      flag% = 0
      For zz = 1 To cmax
         If currprocs(zz) = sp$ Then
            flag% = 1
            Exit For
         End If
      Next
      If flag% Then
         Call CopyStoredProc(sp$, "A")
         numreplaced = numreplaced + 1
      Else
         Call CopyStoredProc(sp$, "C")
         numcopied = numcopied + 1
      End If
   End If
Next
conn.Close
frmMoveStoredProcs.MousePointer = 1
MsgBox Str(numcopied) + " copied and" + Str(numreplaced) + " replaced."


End Sub

Private Sub Command2_Click()

Unload frmMoveStoredProcs

End Sub

Private Sub Form_Load()

cmbServer.Clear
For z = 1 To xmax
   cmbServer.AddItem Xtras(z)
Next
If Server <> "" Then
   cmbServer.Text = Server
Else
   cmbServer.Text = cmbServer.List(0)
End If
List1.Clear
For z = 0 To frmMain!List4.ListCount - 1
   List1.AddItem frmMain!List4.List(z)
Next

End Sub

Public Sub CopyStoredProc(sp$, CorA$)

On Error GoTo errorfound

Set oconn = New adodb.Connection
Set ocmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
oconn.Open strConnect
ocmd1.ActiveConnection = oconn
ocmd1.CommandText = "sp_helptext " + sp$
Set rs = ocmd1.Execute
strText = ""
Do While Not rs.EOF
    strText = strText + rs!Text
    rs.MoveNext
Loop
oconn.Close
If CorA$ = "A" Then strText = Replace(strText, "CREATE PROCEDURE", "ALTER PROCEDURE")
Sql = strText
cmd1.CommandText = Sql
cmd1.Execute
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "Error found - " + Err.Description
   Resume Next
   
End Sub
