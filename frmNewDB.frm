VERSION 5.00
Begin VB.Form frmNewDB 
   Caption         =   "Create Database"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   Icon            =   "frmNewDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "New Database Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then
   MsgBox "Missing New Database Name."
   Exit Sub
End If

On Error GoTo errorfound

Set conn = New adodb.Connection
Set cmd1 = New adodb.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "create database " + Text1.Text
cmd1.CommandText = Sql
cmd1.Execute
conn.Close
LogAction = "Create Database"
LogComment = Text1.Text
Call UpdateLog
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "Error - " + Err.Description
   Resume Next
   
End Sub

Private Sub Command2_Click()

Unload frmNewDB

End Sub

Private Sub Form_Load()

Text1.Text = ""

End Sub
