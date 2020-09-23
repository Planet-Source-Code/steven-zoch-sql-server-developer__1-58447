VERSION 5.00
Begin VB.Form frmAlter 
   Caption         =   "Alter Table"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   Icon            =   "frmAlter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Drop Column"
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   7815
      Begin VB.CommandButton Command4 
         Caption         =   "Drop"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alter Column"
      Height          =   1335
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   7815
      Begin VB.CheckBox Check2 
         Caption         =   "Allow Nulls"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Text            =   "Combo3"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Alter"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Size"
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Column"
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7815
      Begin VB.CheckBox Check1 
         Caption         =   "Allow Nulls"
         Height          =   315
         Left            =   6600
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Text            =   "Combo3"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Size"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmAlter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AlterType As String

Private Sub Combo3_Change()

Select Case Combo3.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Text2.Enabled = True
   Case Else
      Text2.Enabled = False
End Select

End Sub

Private Sub Combo3_Click()

Select Case Combo3.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Text2.Enabled = True
   Case Else
      Text2.Enabled = False
End Select

End Sub

Private Sub Combo4_Change()

Select Case Combo4.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Text3.Enabled = True
   Case Else
      Text3.Enabled = False
End Select

End Sub

Private Sub Combo4_Click()

Select Case Combo4.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Text3.Enabled = True
   Case Else
      Text3.Enabled = False
End Select

End Sub

Private Sub Command1_Click()

If Text1.Text = "" Then
   MsgBox "Enter New Column Name."
   Exit Sub
End If

On Error GoTo errorfound

Set conn = New adodb.Connection
Set cmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Alter table " + TableName + " add " + Text1.Text + " " + Combo3.Text
Select Case Combo3.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Sql = Sql + "(" + Trim(Text2.Text) + ")"
End Select
If Check1.Value Then
   Sql = Sql + " Null"
Else
   Sql = Sql + " Not Null"
End If
cmd1.CommandText = Sql
cmd1.Execute
On Error GoTo 0
LogAction = "Add Column Performed"
LogComment = "New column " + Text1.Text
Call UpdateLog
Text1.Text = ""
Exit Sub

errorfound:
   MsgBox "There was an errror - " + Err.Description
   Resume Next

End Sub

Private Sub Command2_Click()

Unload frmAlter

End Sub

Private Sub Command3_Click()

sysbut% = MsgBox("Are you sure you want to alter the column " + Combo1.Text + "?", 4, "Alter Confirm")
If sysbut% <> vbYes Then Exit Sub

On Error GoTo errorfound

Set conn = New adodb.Connection
Set cmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Alter table " + TableName + " alter column " + Combo1.Text + " " + Combo4.Text
Select Case Combo4.Text
   Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
      Sql = Sql + "(" + Trim(Text3.Text) + ")"
End Select
If Check1.Value Then
   Sql = Sql + " Null"
Else
   Sql = Sql + " Not Null"
End If
cmd1.CommandText = Sql
cmd1.Execute
On Error GoTo 0
Text1.Text = ""
LogAction = "Column Altered"
LogComment = "Column " + Combo1.Text
Call UpdateLog
Exit Sub

errorfound:
   MsgBox "There was an errror - " + Err.Description
   Resume Next

End Sub

Private Sub Command4_Click()

On Error GoTo errorfound

sysbut% = MsgBox("Are you sure you want to drop the column " + Combo2.Text + "?", 4, "Delete Confirm")
If sysbut% <> vbYes Then Exit Sub
Set conn = New adodb.Connection
Set cmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Alter table " + TableName + " drop column " + Combo2.Text
cmd1.CommandText = Sql
cmd1.Execute
On Error GoTo 0
LogAction = "Drop Column Performed"
LogComment = "Column " + Combo2.Text
Call UpdateLog
Exit Sub

errorfound:
   MsgBox "There was an errror - " + Err.Description
   Resume Next

End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Check1.Value = 1
Check2.Value = 1
frmAlter.Caption = "Alter Table " + TableName
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
For z = 0 To frmMain!List3.ListCount - 1
   ci$ = frmMain!List3.List(z)
   If Left(ci$, 6) = "======" Then
      Exit For
   Else
      ci$ = Left(ci$, InStr(ci$, ":") - 1)
      Combo1.AddItem ci$
      Combo2.AddItem ci$
   End If
Next
Combo3.AddItem "bigint"
Combo3.AddItem "binary"
Combo3.AddItem "bit"
Combo3.AddItem "char"
Combo3.AddItem "datetime"
Combo3.AddItem "decimal"
Combo3.AddItem "float"
Combo3.AddItem "image"
Combo3.AddItem "int"
Combo3.AddItem "money"
Combo3.AddItem "nchar"
Combo3.AddItem "ntext"
Combo3.AddItem "numeric"
Combo3.AddItem "nvarchar"
Combo3.AddItem "real"
Combo3.AddItem "smalldatetime"
Combo3.AddItem "smallint"
Combo3.AddItem "smallmoney"
Combo3.AddItem "text"
Combo3.AddItem "timestamp"
Combo3.AddItem "tinyint"
Combo3.AddItem "uniqueidentifier"
Combo3.AddItem "varbinary"
Combo3.AddItem "varchar"
Combo4.AddItem "bigint"
Combo4.AddItem "binary"
Combo4.AddItem "bit"
Combo4.AddItem "char"
Combo4.AddItem "datetime"
Combo4.AddItem "decimal"
Combo4.AddItem "float"
Combo4.AddItem "image"
Combo4.AddItem "int"
Combo4.AddItem "money"
Combo4.AddItem "nchar"
Combo4.AddItem "ntext"
Combo4.AddItem "numeric"
Combo4.AddItem "nvarchar"
Combo4.AddItem "real"
Combo4.AddItem "smalldatetime"
Combo4.AddItem "smallint"
Combo4.AddItem "smallmoney"
Combo4.AddItem "text"
Combo4.AddItem "timestamp"
Combo4.AddItem "tinyint"
Combo4.AddItem "uniqueidentifier"
Combo4.AddItem "varbinary"
Combo4.AddItem "varchar"
Combo1.Text = Combo1.List(0)
Combo2.Text = Combo2.List(0)
Combo3.Text = Combo3.List(0)
Combo4.Text = Combo4.List(0)

End Sub
