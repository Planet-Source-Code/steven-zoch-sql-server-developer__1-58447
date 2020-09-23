VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableCreate 
   Caption         =   "Create Table"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   Icon            =   "frmTableCreate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Drop && Create"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Table"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid tflex 
         Height          =   3495
         Left            =   720
         TabIndex        =   10
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6165
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7335
      Begin VB.CheckBox Check1 
         Caption         =   "Allow Nulls"
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Length"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Column Name"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "New Table Name"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu popci 
         Caption         =   "Create Index"
      End
      Begin VB.Menu poppk 
         Caption         =   "Create Primary Key"
      End
      Begin VB.Menu popri 
         Caption         =   "Remove Index"
      End
      Begin VB.Menu poprk 
         Caption         =   "Remove Primary Key"
      End
      Begin VB.Menu popdiv 
         Caption         =   "-"
      End
      Begin VB.Menu popdc 
         Caption         =   "Delete Column"
      End
      Begin VB.Menu popex 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTableCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cmax As Integer
Public CRow As Integer
Public CCol As Integer

Private Sub Combo1_Click()

Select Case Combo1.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Text2.Enabled = True
   Case Else
      Text2.Enabled = False
End Select

End Sub

Private Sub Command1_Click()

If Text3.Text = "" Then
   MsgBox "Missing Column Name"
   Exit Sub
End If

cmax = cmax + 1
ColData(cmax, 1) = Text3.Text
ColData(cmax, 2) = Combo1.Text
ColData(cmax, 3) = Text2.Text
If Check1.Value Then
   ColData(cmax, 4) = "Y"
Else
   ColData(cmax, 4) = "N"
End If

tflex.Rows = cmax + 1
tflex.Row = cmax
For z = 1 To 4
   tflex.Col = z
   tflex.Text = ColData(cmax, z)
Next

Text3.Text = ""
Text2.Text = ""
Text3.SetFocus

End Sub

Private Sub Command2_Click()

Unload frmTableCreate

End Sub

Private Sub Command3_Click()

If Text1.Text = "" Then
   MsgBox "No Table Name Given"
   Exit Sub
End If

LogAction = "Table Created"
LogComment = "new table " + Text1.Text + " in database " + DBName
CSql = "Create Table dbo." + Text1.Text + " (" + vbCrLf
For z = 1 To cmax
   CSql = CSql + "[" + ColData(z, 1) + "] [" + ColData(z, 2) + "] "
   Select Case ColData(z, 2)
      Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
         CSql = CSql + "(" + Trim(ColData(z, 3)) + ") "
   End Select
   If ColData(z, 4) = "Y" Then
      CSql = CSql + "NULL"
   Else
      CSql = CSql + "NOT NULL"
   End If
   CSql = CSql + "," + vbCrLf
Next
CSql = Left(CSql, Len(CSql) - 2) + vbCrLf + ") ON [PRIMARY]" + vbCrLf

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = CSql
cmd1.Execute

For z = 1 To tflex.Rows - 1
   tflex.Row = z
   tflex.Col = 0
   If tflex.Text = "P" Then
      tflex.Col = 1
      acol$ = tflex.Text
      Sql = "alter table " + Text1.Text + " add constraint PK_" + acol$ + " primary key (" + acol$ + ")"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
   If tflex.Text = "I" Then
      tflex.Col = 1
      acol$ = tflex.Text
      Sql = "create index INDX_" + acol$ + " on " + Text1.Text + " (" + acol$ + ")"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
Next

conn.Close
cmax = 0
On Error GoTo 0
Call UpdateLog
Form_Load
Text1.SetFocus
Exit Sub

errorfound:
   MsgBox "Error found - " + Err.Description
   Resume Next
   
End Sub

Private Sub Command4_Click()

If Text1.Text = "" Then
   MsgBox "No Table Name Given"
   Exit Sub
End If

LogAction = "Table Recreated"
LogComment = "recreated table " + Text1.Text + " in database " + DBName
CSql = "Create Table dbo." + Text1.Text + " (" + vbCrLf
For z = 1 To cmax
   CSql = CSql + "[" + ColData(z, 1) + "] [" + ColData(z, 2) + "] "
   Select Case ColData(z, 2)
      Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
         CSql = CSql + "(" + Trim(ColData(z, 3)) + ") "
   End Select
   If ColData(z, 4) = "Y" Then
      CSql = CSql + "NULL"
   Else
      CSql = CSql + "NOT NULL"
   End If
   CSql = CSql + "," + vbCrLf
Next
CSql = Left(CSql, Len(CSql) - 2) + vbCrLf + ") ON [PRIMARY]" + vbCrLf

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

cmd1.CommandText = "Drop table " + Text1.Text
cmd1.Execute

cmd1.CommandText = CSql
cmd1.Execute

For z = 1 To tflex.Rows - 1
   tflex.Row = z
   tflex.Col = 0
   If tflex.Text = "P" Then
      tflex.Col = 1
      acol$ = tflex.Text
      Sql = "alter table " + Text1.Text + " add constraint PK_" + acol$ + " primary key (" + acol$ + ")"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
   If tflex.Text = "I" Then
      tflex.Col = 1
      acol$ = tflex.Text
      Sql = "create index INDX_" + acol$ + " on " + Text1.Text + " (" + acol$ + ")"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
Next

conn.Close
On Error GoTo 0
cmax = 0
Call UpdateLog
Form_Load
Exit Sub

errorfound:
   Resume Next

End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text2.Enabled = False
Combo1.Clear
Combo1.AddItem "bigint"
Combo1.AddItem "binary"
Combo1.AddItem "bit"
Combo1.AddItem "char"
Combo1.AddItem "datetime"
Combo1.AddItem "decimal"
Combo1.AddItem "float"
Combo1.AddItem "image"
Combo1.AddItem "int"
Combo1.AddItem "money"
Combo1.AddItem "nchar"
Combo1.AddItem "ntext"
Combo1.AddItem "numeric"
Combo1.AddItem "nvarchar"
Combo1.AddItem "real"
Combo1.AddItem "smalldatetime"
Combo1.AddItem "smallint"
Combo1.AddItem "smallmoney"
Combo1.AddItem "text"
Combo1.AddItem "timestamp"
Combo1.AddItem "tinyint"
Combo1.AddItem "uniqueidentifier"
Combo1.AddItem "varbinary"
Combo1.AddItem "varchar"
Combo1.Text = Combo1.List(0)

Check1.Value = 1
tflex.Cols = 5
tflex.ColWidth(0) = 200
tflex.ColWidth(1) = 2200
tflex.ColWidth(2) = 1500
tflex.ColWidth(3) = 1000
tflex.ColWidth(4) = 800
tflex.Rows = 2
tflex.Row = 0
tflex.Col = 1
tflex.Text = "Column"
tflex.Col = 2
tflex.Text = "Type"
tflex.Col = 3
tflex.Text = "Length"
tflex.Col = 4
tflex.Text = "Nulls"
tflex.Rows = 1

End Sub

Private Sub popci_Click()

tflex.Row = CRow
tflex.Col = 0
tflex.Text = "I"

End Sub

Private Sub poppk_Click()

tflex.Row = CRow
tflex.Col = 0
tflex.Text = "P"

End Sub

Private Sub popri_Click()

tflex.Row = CRow
tflex.Col = 0
tflex.Text = ""

End Sub

Private Sub poprk_Click()

tflex.Row = CRow
tflex.Col = 0
tflex.Text = ""

End Sub

Private Sub tflex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  CRow = tflex.Row
  CCol = tflex.Col
  If Button = 2 Then
     PopupMenu mnuPop
  End If
  
End Sub

Public Sub popdc_click()

   If CRow = 0 Then Exit Sub
   tflex.Row = CRow
   tflex.Col = 1
   SelectedCol = tflex.Text
   If SelectedCol = "" Then Exit Sub
   
   sysbut% = MsgBox("Are you sure you want to delete the Column " + SelectedCol + "?", 4, "Delete Confirm")
   If sysbut% <> vbYes Then Exit Sub
   
   For z = 1 To cmax
     tflex.Row = z
     For zz = 1 To 4
        tflex.Col = zz
        tflex.Text = ""
     Next
   Next
   
   For z = 1 To cmax
      If ColData(z, 1) = SelectedCol Then
         pos = z
         Exit For
      End If
   Next
   
   For z = pos To cmax - 1
     For zz = 1 To 4
         ColData(z, zz) = ColData(z + 1, zz)
     Next
   Next
   cmax = cmax - 1
   
   tflex.Rows = cmax + 1
   For z = 1 To cmax
      tflex.Row = z
      For zz = 1 To 4
         tflex.Col = zz
         tflex.Text = ColData(z, zz)
      Next
   Next
   
End Sub
Public Sub popex_click()

   mnuPop.Visible = False
   
End Sub

