VERSION 5.00
Begin VB.Form frmSearchMain 
   Caption         =   "Search Main"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   Icon            =   "frmSearchMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   4815
      Begin VB.OptionButton Option9 
         Caption         =   "Table Only"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Entire Databases"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Current DB"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      Begin VB.OptionButton Option6 
         Caption         =   "Data Str in Triggers"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Data Str in Stored Proc"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Data Str in Data (long process)"
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Data Str in Columns"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Data Type"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Columns"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Search String"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Note:  Some searching might require a lot of time..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmSearchMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

SearchDone = False
SearchInProgress = False

If Text1.Text = "" Then
   MsgBox "Missing Search String."
   Exit Sub
End If

If Option1.Value Then
   SearchType = "col"
End If
If Option2.Value Then
   SearchType = "dt"
End If
If Option3.Value Then
   SearchType = "dsc"
End If
If Option4.Value Then
   SearchType = "dsd"
End If
If Option5.Value Then
   SearchType = "dss"
End If
If Option6.Value Then
   SearchType = "dst"
End If
If Option7.Value Then
   SearchType = "default " + SearchType
End If
If Option8.Value Then
   SearchType = "entire " + SearchType
End If
If Option9.Value Then
   SearchType = "table " + SearchType
End If

SearchStr = Text1.Text
frmSearch.Show

End Sub

Private Sub Command2_Click()

SearchInProgress = False
Unload frmSearchMain

End Sub

Private Sub Form_Activate()

Text1.SetFocus

End Sub

Private Sub Form_Load()

Text1.Text = ""
Option4.Value = 1
If DBName <> "" Then
   Option7.Value = 1
Else
   Option8.Value = 1
   Option7.Enabled = False
End If

End Sub

Private Sub Option7_Click()

If Option7.Value Then
   If DBName = "" Then
      MsgBox "A Database must first be selected before this feature can be performed."
      Option7.Value = 0
      Option8.Value = 1
   End If
End If

If Option7.Value Then
      If frmMain!List4.ListCount = 0 Then
         Option5.Enabled = False
      Else
         Option5.Enabled = True
      End If
      If frmMain!List5.ListCount = 0 Then
         Option6.Enabled = False
      Else
         Option6.Enabled = True
      End If
End If

End Sub

Private Sub Option8_Click()

If Option8.Value Then
   Option5.Enabled = True
   Option6.Enabled = True
End If
End Sub
