VERSION 5.00
Begin VB.Form frmRename 
   Caption         =   "Rename"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rename"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "New Name"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Old Name"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton Option4 
         Caption         =   "Trigger"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Stored Procedure"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Column"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Table"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RenameType As String

Private Sub Command1_Click()

If Text1.Text = "" Then
  MsgBox "Need to provide new name."
  Exit Sub
End If

On Error GoTo errorfound

Set conn = New adodb.Connection
Set cmd1 = New adodb.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Select Case RenameType
   Case "Table", "Procedure", "Trigger"
      Sql = "sp_rename '" + Combo1.Text + "', '" + Text1.Text + "'"
   Case "Column"
      Sql = "sp_rename '" + TableName + "." + Combo1.Text + "', '" + Text1.Text + "'"

End Select
cmd1.CommandText = Sql
cmd1.Execute
conn.Close
On Error GoTo 0
LogAction = "Rename " + RenameType
LogComment = Combo1.Text + " (old) " + Text1.Text + " (new)"
Call UpdateLog
Exit Sub

errorfound:
   MsgBox "There was the following error - " + Err.Description
   Resume Next
   
End Sub

Private Sub Command2_Click()

Unload frmRename

End Sub

Private Sub Form_Load()

Combo1.Clear
Text1.Text = ""
If TableName = "" Then Option2.Enabled = False
If frmMain!List4.ListCount = 0 Then Option3.Enabled = False
If frmMain!List5.ListCount = 0 Then Option4.Enabled = False

End Sub

Private Sub Option1_Click()

Combo1.Clear
For z = 0 To frmMain!List2.ListCount - 1
   Combo1.AddItem frmMain!List2.List(z)
Next
Combo1.Text = Combo1.List(0)
RenameType = "Table"

End Sub

Private Sub Option2_Click()

Combo1.Clear
For z = 0 To frmMain!List3.ListCount - 1
   ci$ = frmMain!List3.List(z)
   ci$ = Left(frmMain!List3.List(z), InStr(frmMain!List3.List(z), ":") - 1)
   Combo1.AddItem ci$
Next
Combo1.Text = Combo1.List(0)
RenameType = "Column"

End Sub

Private Sub Option3_Click()

Combo1.Clear
For z = 0 To frmMain!List4.ListCount - 1
   Combo1.AddItem frmMain!List4.List(z)
Next
Combo1.Text = Combo1.List(0)
RenameType = "Procedure"

End Sub

Private Sub Option4_Click()

Combo1.Clear
For z = 0 To frmMain!List5.ListCount - 1
   Combo1.AddItem frmMain!List5.List(z)
Next
Combo1.Text = Combo1.List(0)
RenameType = "Trigger"

End Sub
