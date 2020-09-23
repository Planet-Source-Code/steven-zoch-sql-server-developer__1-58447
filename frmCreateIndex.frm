VERSION 5.00
Begin VB.Form frmCreateIndex 
   Caption         =   "Create Index"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "frmCreateIndex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Nonclustered"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clustered"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Unique"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "New Index Name"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCreateIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      List2.AddItem List1.List(z)
      Exit For
   End If
Next

End Sub

Private Sub Command3_Click()

Unload frmCreateIndex

End Sub

Private Sub Command4_Click()

If Text1.Text = "" Then
   MsgBox "Missing New Index Name."
   Exit Sub
End If

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "Create "
If Check1.Value Then Sql = Sql + "unique "
If Option1.Value Then Sql = Sql + "clustered "
If Option2.Value Then Sql = Sql + "nonclustered "
Sql = Sql + "Index " + Text1.Text + " on " + TableName + " ("
Colstr = ""
For z = 0 To List2.ListCount - 1
   Colstr = Colstr + List2.List(z) + ", "
Next
Colstr = Left(Colstr, Len(Colstr) - 2)
Sql = Sql + Colstr + ")"
cmd1.CommandText = Sql
cmd1.Execute
If Err.Number = 0 Then
   Text1.Text = ""
   List2.Clear
   keystr = ""
End If
On Error GoTo 0
Exit Sub

errorfound:
  MsgBox "There was the following error - " + Err.Description
  Resume Next

End Sub

Private Sub Form_Load()

Text1.Text = ""
List1.Clear
List2.Clear
keystr = ""
For z = 0 To frmMain!List3.ListCount - 1
   ci$ = frmMain!List3.List(z)
   ci$ = Left(ci$, InStr(ci$, ":") - 1)
   List1.AddItem ci$
Next
Check1.Value = 1
Option1.Value = 1

End Sub
