VERSION 5.00
Begin VB.Form frmDelete 
   Caption         =   "Delete"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmDelete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   120
         Width           =   495
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
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DeleteType As String

Private Sub Command1_Click()

sysbut% = MsgBox("Are you sure you want to delete the " + DeleteType + "s?", 4, "Delete Confirm")
If sysbut% <> vbYes Then Exit Sub

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      Select Case DeleteType
         Case "Table", "Procedure", "Trigger"
            Sql = "drop " + DeleteType + " " + List1.List(z)
         Case "Column"
            Sql = "Alter table " + TableName + " drop column " + List1.List(z)
      End Select
      cmd1.CommandText = Sql
      cstate$ = Sql
      msg$ = ""
      If DeleteType <> "Column" Then
         dcheck& = CheckDependencies(cstate$, msg$)
         If dcheck& Then
            sysbut% = MsgBox("The following dependencies are connected to this object:" + vbCrLf + msg$ + vbCrLf + "Are you sure you want to drop it?", 4, "Drop confirm")
            If sysbut% <> vbYes Then GoTo Done
         End If
      End If
      cmd1.Execute
   End If
Done:
Next
conn.Close
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      LogAction = DeleteType + " Deleted"
      LogComment = DeleteType + " " + List1.List(z) + " in database " + DBName
      Call UpdateLog
      List1.RemoveItem z
   End If
Next
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "There was the following error - " + Err.Description
   Resume Next
   
End Sub

Private Sub Command2_Click()

If DeleteType = "Table" And DeletePerformed And TableName <> "" Then
   ReloadNeeded = True
End If
Unload frmDelete

End Sub

Private Sub Form_Load()

List1.Clear
If TableName = "" Then Option2.Enabled = False
If frmMain!List4.ListCount = 0 Then Option3.Enabled = False
If frmMain!List5.ListCount = 0 Then Option4.Enabled = False

End Sub

Private Sub Option1_Click()

List1.Clear
For z = 0 To frmMain!List2.ListCount - 1
   List1.AddItem frmMain!List2.List(z)
Next
DeleteType = "Table"

End Sub

Private Sub Option2_Click()

List1.Clear
For z = 0 To frmMain!List3.ListCount - 1
   ci$ = frmMain!List3.List(z)
   ci$ = Left(frmMain!List3.List(z), InStr(frmMain!List3.List(z), ":") - 1)
   List1.AddItem ci$
Next
DeleteType = "Column"

End Sub

Private Sub Option3_Click()

List1.Clear
For z = 0 To frmMain!List4.ListCount - 1
   List1.AddItem frmMain!List4.List(z)
Next
DeleteType = "Procedure"

End Sub

Private Sub Option4_Click()

List1.Clear
For z = 0 To frmMain!List5.ListCount - 1
   List1.AddItem frmMain!List5.List(z)
Next
DeleteType = "Trigger"

End Sub
