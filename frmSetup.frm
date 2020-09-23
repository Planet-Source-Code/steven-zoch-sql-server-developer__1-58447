VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup SQLDeveloper"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   6615
      Begin VB.CommandButton Command3 
         Caption         =   "Finish"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "sa"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Password"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Username"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   $"frmSetup.frx":0442
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "Exit Setup"
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   1230
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "New Database"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   $"frmSetup.frx":04ED
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmSetup.frx":0611
         Top             =   720
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Add any additional Servers or IP Addresses to include.  You can add or delete to this list under File/Modify Servers."
         Height          =   615
         Left            =   3120
         TabIndex        =   5
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Select the initial Sql Server for the first run."
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ServerName As String

Private Sub Command1_Click()

frmSetup.MousePointer = 11
DBLog = ""
For z = 0 To List2.ListCount - 1
   If List2.Selected(z) Then
      DBLog = List2.List(z)
      Exit For
   End If
Next

If DBLog = "" And Text2.Text = "" Then
   frmSetup.MousePointer = 1
   MsgBox "Select either a defined database or a new one."
   Exit Sub
End If

   Open "c:\SqlDeveloperOptions.ini" For Output As #1
   Print #1, ServerName
   Print #1, txtUser.Text
   Print #1, txtPassword.Text
   Print #1, "&H00FFFFFF"
   Print #1, "&H00800000"
   Print #1, "MS Sans Serif"
   Print #1, "8"
   Print #1, "N"
   Print #1, "&H80000005"
   Print #1, "&H80000008"
   Print #1, "&H80000005"
   Print #1, "&H80000008"
   Print #1, "&H80000005"
   Print #1, "&H80000008"
   Print #1, "&H80000005"
   Print #1, "&H80000008"
   Print #1, "&H80000005"
   Print #1, "&H80000008"
   
   xmax = 0
   Erase Xtras
   If Text1.Text = "" Then
      For z = 0 To List1.ListCount - 1
         xmax = xmax + 1
         Xtras(xmax) = List1.List(z)
      Next
      GoTo ReadyToCreateTable
   End If
   x$ = ""
   For z = 1 To Len(Text1.Text)
      c$ = Mid(Text1.Text, z, 1)
      Select Case c$
         Case Chr(13)
            xmax = xmax + 1
            Xtras(xmax) = x$
            x$ = ""
         Case Chr(10)
         Case Else
            x$ = x$ + c$
      End Select
   Next
'new create LogTable
ReadyToCreateTable:
   Print #1, Str(xmax)
   For z = 1 To xmax
      Print #1, Xtras(z)
   Next

  ServerTableCreateString = "Create Table dbo.SqlDeveloperServers ([ServerName] [varchar] (50) NULL, [ServerUser] [varchar] (50) NULL, [ServerPassword] [varchar] (50) NULL) ON [PRIMARY]"
  TableCreateString = "Create Table dbo.SqlDeveloperLog ([LogAction] [varchar] (50) NULL, [LogUser] [varchar] (50) NULL, [LogDate] [datetime]  NULL, [LogComment] [varchar] (255) NOT NULL) ON [PRIMARY]"
   flag% = 0
   For z = 0 To List2.ListCount - 1
      If List2.Selected(z) Then
         DBLog = List2.List(z)
         flag% = 1
      End If
   Next
   If flag% = 0 Then
      DBLog = Text2.Text
   End If

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
On Error Resume Next

If flag% = 0 Then
   strConnect = "Provider=SQLOLEDB;server=" + ServerName + ";uid=" + txtUser.Text + ";pwd=" + txtPassword.Text + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   Sql = "Create database " + DBLog
   cmd1.CommandText = Sql
   cmd1.Execute
   conn.Close
End If

strConnect = "Provider=SQLOLEDB;server=" + ServerName + ";database=" + DBLog + ";uid=" + txtUser.Text + ";pwd=" + txtPassword.Text + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = TableCreateString
cmd1.Execute
cmd1.CommandText = ServerTableCreateString
cmd1.Execute
For z = 1 To xmax
   Sql = "Select ServerName from SqlDeveloperServers where ServerName='" + Xtras(z) + "'"
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.EOF Or rs.BOF Then
      Sql = "insert into SqlDeveloperServers (ServerName, ServerUser, ServerPassword) values ('" + Xtras(z) + "', '" + txtUser.Text + "', '" + txtPassword.Text + "')"
      cmd1.CommandText = Sql
      cmd1.Execute
   End If
Next
conn.Close

Done:
   LogServer = ServerName
   Print #1, LogServer
   Print #1, DBLog
   frmSetup.MousePointer = 1
   Close
   On Error GoTo 0
   Unload frmSetup
   
End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Command3_Click()

If txtUser.Text = "" Or txtPassword.Text = "" Then
   MsgBox "User and Password required."
   Exit Sub
End If

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + ServerName + ";uid=" + txtUser.Text + ";pwd=" + txtPassword.Text + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "sp_databases"
cmd1.CommandText = Sql
Set rs = cmd1.Execute

Do While Not rs.EOF
   List2.AddItem rs!Database_Name
   rs.MoveNext
Loop
conn.Close
Frame2.Visible = True

End Sub

Private Sub Form_Load()

Dim SQLApp As SQLDMO.Application
Dim Names As SQLDMO.NameList
Dim indx As Integer

List1.Clear
Frame2.Visible = False
Frame3.Visible = False
Text1.Text = ""
Text2.Text = ""

Set SQLApp = New SQLDMO.Application
Set Names = SQLApp.ListAvailableSQLServers
For indx = 1 To Names.Count
   List1.AddItem Names.Item(indx)
Next

Set Names = Nothing
Set SQLApp = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Open "c:\SqlDeveloperOptions.ini" For Random As #1 Len = 1
l& = LOF(1)
Close
If l& = 0 Then
   End
End If

End Sub

Private Sub List1_Click()

For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      ServerName = List1.List(z)
      Exit For
   End If
Next
Frame3.Visible = True

End Sub
