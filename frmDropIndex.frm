VERSION 5.00
Begin VB.Form frmDropIndex 
   Caption         =   "Drop Index"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   Icon            =   "frmDropIndex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Drop"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Index to Drop"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmDropIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HasIndex As Boolean

Private Sub Command1_Click()

Unload frmDropIndex

End Sub

Private Sub Command2_Click()

sysbut% = MsgBox("Are you sure you want to drop the index " + Combo1.Text + "?", 4, "Drop Confirm")
If sysbut% <> vbYes Then Exit Sub

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Drop Index " + TableName + "." + Combo1.Text
cmd1.CommandText = Sql
cmd1.Execute
conn.Close
Form_Load

End Sub

Private Sub Form_Activate()

If HasIndex Then Exit Sub

Unload frmDropIndex

End Sub

Private Sub Form_Load()

Combo1.Clear
HasIndex = True

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "sp_helpindex " + TableName
cmd1.CommandText = Sql
Set rs = cmd1.Execute
If rs.EOF Or rs.BOF Then
   MsgBox "No Indexes for Table " + TableName
   HasIndex = False
   Exit Sub
End If

Do While Not rs.EOF
   Combo1.AddItem rs!Name
   rs.MoveNext
Loop
conn.Close

End Sub
