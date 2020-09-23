VERSION 5.00
Begin VB.Form frmModServers 
   Caption         =   "Modify Servers"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   Icon            =   "frmModServers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command6 
         Caption         =   "Test"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   3270
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Remove"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3270
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<=="
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   3270
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "New or Additional Servers"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Current Servers in Use"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmModServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Open "c:\SqlDeveloperOptions.ini" For Output As #1
   Print #1, Server
   Print #1, UID
   Print #1, PWD
   Print #1, Str(TForeColor)
   Print #1, Str(TBackColor)
   Print #1, TFontName
   Print #1, Str(TFontSize)
   Print #1, NullReq
   Print #1, Str(List1TBackColor)
   Print #1, Str(List1TForeColor)
   Print #1, Str(List2TBackColor)
   Print #1, Str(List2TForeColor)
   Print #1, Str(List3TBackColor)
   Print #1, Str(List3TForeColor)
   Print #1, Str(List4TBackColor)
   Print #1, Str(List4TForeColor)
   Print #1, Str(List5TBackColor)
   Print #1, Str(List5TForeColor)
   Print #1, Str(List1.ListCount)
   xmax = List1.ListCount
   Erase Xtras
   For z = 0 To List1.ListCount - 1
      Print #1, List1.List(z)
      Xtras(z + 1) = List1.List(z)
   Next
   Print #1, LogServer
   Print #1, DBLog
   Close
   
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 1 To xmax
   Sql = "Select ServerName from SqlDeveloperServers where servername='" + Xtras(z) + "'"
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      Sql = "insert into SqlDeveloperServers (ServerName, ServerUser, ServerPassword) values ('" + Xtras(z) + "', '" + UID + "', '" + PWD + "')"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
Next
conn.Close
   
   
   Unload frmModServers
   
End Sub

Private Sub Command2_Click()

Unload frmModServers

End Sub

Private Sub Command3_Click()

If Text1.Text = "" Then
   MsgBox "Enter Server or IP Address"
   Exit Sub
End If
List1.AddItem Text1.Text

End Sub

Private Sub Command4_Click()

For z = 0 To List2.ListCount - 1
   If List2.Selected(z) Then
      List1.AddItem List2.List(z)
      List2.RemoveItem z
      Exit For
   End If
Next
List2.Refresh
If List2.ListCount = 0 Then Command4.Visible = False

End Sub

Private Sub Command5_Click()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      List2.AddItem List1.List(z)
      Sql = "Delete from SqlDeveloperServers where servername='" + List1.List(z) + "'"
      cmd1.CommandText = Sql
      cmd1.Execute
      List1.RemoveItem z
      Exit For
   End If
Next
conn.Close
List1.Refresh

End Sub

Private Sub Command6_Click()

If Text1.Text = "" Then
   MsgBox "Enter Server or IP Address"
   Exit Sub
End If

On Error GoTo errorfound

Tested = True
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Text1.Text + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
conn.Close
If Tested = False Then
   MsgBox "The Connection Failed."
Else
   MsgBox "The Connection Was Successful."
End If
conn.Close
On Error GoTo 0
Exit Sub

errorfound:
   Tested = False
   Resume Next

End Sub

Private Sub Form_Load()

Dim SQLApp As SQLDMO.Application
Dim Names As SQLDMO.NameList
Dim indx As Integer

List1.Clear
List2.Clear
Text1.Text = ""

Open "c:\SqlDeveloperOptions.ini" For Input As #1
For z = 1 To 18
   Line Input #1, d$
Next
Line Input #1, d$
For z = 1 To Val(d$)
   Line Input #1, d$
   List1.AddItem d$
Next
Close

Set SQLApp = New SQLDMO.Application
Set Names = SQLApp.ListAvailableSQLServers
For indx = 1 To Names.Count
   flag% = 0
   For z = 0 To List1.ListCount - 1
      If List1.List(z) = Names.Item(indx) Then
         flag% = 1
         Exit For
      End If
   Next
   If flag% = 0 Then List2.AddItem Names.Item(indx)
Next
If List2.ListCount = 0 Then
   Command4.Visible = False
End If

Set Names = Nothing
Set SQLApp = Nothing

End Sub
