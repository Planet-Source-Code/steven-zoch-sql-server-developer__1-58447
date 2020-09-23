VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "SQL Statement"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute (F5)"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmText.frx":030A
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If SearchInProgress = False Then frmMain!Text1.Text = Text1.Text
Unload frmText

End Sub

Private Sub Command2_Click()

On Error GoTo errorfound

frmText.MousePointer = 11
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = Text1.Text
If Left(LCase(Text1.Text), 7) = "select " Then
   Set rs = cmd1.Execute
   MsgBox "Recordset was returned, but cannot be displayed here."
Else
   cmd1.Execute
End If
frmText.MousePointer = 1
conn.Close
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "Error occurred - " + Err.Description
   Resume Next
   
End Sub

Private Sub Command3_Click()

Printer.Print Text1.Text
Printer.EndDoc

End Sub

Private Sub Form_Load()

Text1.ForeColor = TForeColor
Text1.BackColor = TBackColor
Text1.FontName = TFontName
Text1.FontSize = TFontSize
Text1.Text = frmMain!Text1.Text
If SearchInProgress Then
   Command2.Enabled = False
End If
frmText.Caption = "SQL Statement"
If SPName <> "" Then frmText.Caption = "Stored Procedure " + SPName
If TRName <> "" Then frmText.Caption = "Trigger " + TRName

End Sub

Private Sub Form_Resize()

On Error Resume Next
Command1.Top = frmText.Height - 990
Command2.Top = frmText.Height - 990
Command3.Top = frmText.Height - 990
Text1.Height = frmText.Height - 1215
Text1.Width = frmText.Width - 50
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

If SearchInProgress = False Then frmMain!Text1.Text = Text1.Text
SPName = ""
TRName = ""

End Sub

Private Sub Text1_Change()

If InStr(Text1.Text, "CREATE PROCEDURE") Then
    Text1.Text = Replace(Text1.Text, "CREATE PROCEDURE", "ALTER PROCEDURE")
End If
If InStr(Text1.Text, "CREATE TRIGGER") Then
    Text1.Text = Replace(Text1.Text, "CREATE TRIGGER", "ALTER TRIGGER")
End If
If InStr(Text1.Text, "create procedure") Then
    Text1.Text = Replace(Text1.Text, "create procedure", "alter procedure")
End If
If InStr(Text1.Text, "create trigger") Then
    Text1.Text = Replace(Text1.Text, "create trigger", "alter trigger")
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
   Case 116
       Command2_Click
   Case 27
       Command1_Click
End Select

End Sub
