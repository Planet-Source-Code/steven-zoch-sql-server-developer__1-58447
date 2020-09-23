VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Sql Window"
      Height          =   2535
      Left            =   3720
      TabIndex        =   10
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command7 
         Caption         =   "More..."
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   480
         TabIndex        =   14
         Text            =   "This is a test."
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back Color..."
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Fore Color..."
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Font..."
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Nulls as Data"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPWD 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmbServer 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tested As Boolean

Private Sub cmbServer_Click()

For z = 0 To cmbServer.ListCount - 1
   If cmbServer.Text = cmbServer.List(z) Then
      txtUID.Text = List1.List(z)
      txtPWD.Text = List2.List(z)
      Exit For
   End If
Next

End Sub

Private Sub Command1_Click()

If cmbServer.Text = "" Or txtUID.Text = "" Or txtPWD.Text = "" Then
   MsgBox "Missing Data.  Make sure all fields are defined and tested."
   Exit Sub
End If

Open "c:\SqlDeveloperOptions.ini" For Output As #1
Print #1, cmbServer.Text
Print #1, txtUID.Text
Print #1, txtPWD.Text
Print #1, Str(TForeColor)
Print #1, Str(TBackColor)
Print #1, TFontName
Print #1, Str(TFontSize)
If Check1.Value Then
   Print #1, "Y"
Else
   Print #1, "N"
End If
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
frmMain!List1.BackColor = Str(List1TBackColor)
frmMain!List1.ForeColor = Str(List1TForeColor)
frmMain!List2.BackColor = Str(List2TBackColor)
frmMain!List2.ForeColor = Str(List2TForeColor)
frmMain!List3.BackColor = Str(List3TBackColor)
frmMain!List3.ForeColor = Str(List3TForeColor)
frmMain!List4.BackColor = Str(List4TBackColor)
frmMain!List4.ForeColor = Str(List4TForeColor)
frmMain!List5.BackColor = Str(List5TBackColor)
frmMain!List5.ForeColor = Str(List5TForeColor)
Print #1, Str(xmax)
For z = 1 To xmax
   Print #1, Xtras(z)
Next
Print #1, LogServer
Print #1, DBLog

Server = cmbServer.Text
UID = txtUID.Text
PWD = txtPWD.Text
Close

Unload frmOptions

End Sub

Private Sub Command2_Click()

Unload frmOptions

End Sub

Private Sub Command3_Click()

On Error GoTo errorfound

Tested = True

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + cmbServer.Text + ";uid=" + txtUID.Text + ";pwd=" + txtPWD.Text + ";"
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

Private Sub Command4_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   Text1.BackColor = CommonDialog1.Color
   TBackColor = CommonDialog1.Color
End If

End Sub

Private Sub Command5_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   Text1.ForeColor = CommonDialog1.Color
   TForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command6_Click()

CommonDialog1.Flags = 1
CommonDialog1.ShowFont
If CommonDialog1.FontName <> "" Then
   Text1.FontName = CommonDialog1.FontName
   TFontName = CommonDialog1.FontName
   Text1.FontSize = CommonDialog1.FontSize
   TFontSize = CommonDialog1.FontSize
End If

End Sub

Private Sub Command7_Click()

frmOptions2.Show 1

End Sub

Private Sub Form_Load()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

cmbServer.Clear
For z = 1 To xmax
   cmbServer.AddItem Xtras(z)
   Sql = "Select * from SqlDeveloperServers where servername='" + Xtras(z) + "'"
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   List1.AddItem rs!ServerUser
   List2.AddItem rs!ServerPassword
Next
If Server <> "" Then
   cmbServer.Text = Server
Else
   cmbServer.Text = cmbServer.List(0)
End If
conn.Close

Tested = False
Open "c:\SqlDeveloperOptions.ini" For Random As #1 Len = 1
l& = LOF(1)
Close
If l& Then
   Open "c:\SqlDeveloperOptions.ini" For Input As #1
   Line Input #1, Server
   Line Input #1, UID
   Line Input #1, PWD
   Line Input #1, TForeColor
   Line Input #1, TBackColor
   Line Input #1, TFontName
   Line Input #1, TFontSize
   Line Input #1, NullReq
   Line Input #1, List1TBackColor
   Line Input #1, List1TForeColor
   Line Input #1, List2TBackColor
   Line Input #1, List2TForeColor
   Line Input #1, List3TBackColor
   Line Input #1, List3TForeColor
   Line Input #1, List4TBackColor
   Line Input #1, List4TForeColor
   Line Input #1, List5TBackColor
   Line Input #1, List5TForeColor
   Line Input #1, x$
   xmax = Val(x$)
   Erase Xtras
   For z = 1 To xmax
      Line Input #1, Xtras(z)
   Next
   Line Input #1, LogServer
   Line Input #1, DBLog
   Close
End If

frmMain!List1.BackColor = Str(List1TBackColor)
frmMain!List2.BackColor = Str(List2TBackColor)
frmMain!List3.BackColor = Str(List3TBackColor)
frmMain!List4.BackColor = Str(List4TBackColor)
frmMain!List5.BackColor = Str(List5TBackColor)
frmMain!List1.ForeColor = Str(List1TForeColor)
frmMain!List2.ForeColor = Str(Lirst2TForeColor)
frmMain!List3.ForeColor = Str(List3TForeColor)
frmMain!List4.ForeColor = Str(List4TForeColor)
frmMain!List5.ForeColor = Str(List5TForeColor)

Text1.BackColor = TBackColor
Text1.ForeColor = TForeColor
Text1.FontName = TFontName
Text1.FontSize = Val(TFontSize)
If NullReq = "Y" Then
   Check1.Value = 1
Else
   Check1.Value = 0
End If
cmbServer.Text = Server
txtUID.Text = UID
txtPWD.Text = PWD

End Sub

