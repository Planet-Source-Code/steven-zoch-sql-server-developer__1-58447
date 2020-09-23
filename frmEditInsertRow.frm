VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEditInsertRow 
   Caption         =   "Edit Row"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmEditInsertRow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   13680
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update/Insert"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   1720
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmEditInsertRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim olddata As String
Dim ChangeInProgress As Boolean

Private Sub Command1_Click()

Select Case EditOrInsert
   Case "Edit"
      Call UpdateRow
   Case "Insert"
      Call InsertRow
End Select

End Sub

Private Sub Command2_Click()

Unload frmEditInsertRow

End Sub

Private Sub Command3_Click()

For z = 1 To flex.Cols - 1
   If flex.ColWidth(z) - 50 > 0 Then
      flex.ColWidth(z) = flex.ColWidth(z) - 50
   End If
Next

End Sub

Private Sub Command4_Click()

For z = 1 To flex.Cols - 1
   flex.ColWidth(z) = flex.ColWidth(z) + 50
Next

End Sub

Private Sub flex_Click()

If flex.MouseRow = 0 Then Exit Sub
olddata = flex.Text
Text1.Text = olddata
Text1.SelStart = 0
Text1.Height = flex.CellHeight
Text1.Width = flex.CellWidth
Text1.Move flex.CellLeft + flex.Left, flex.CellTop + flex.Top, flex.CellWidth, flex.CellHeight
Text1.Visible = True
Text1.SetFocus
ChangeInProgress = True

End Sub

Private Sub flex_LeaveCell()

If Text1.Text = "" And ChangeInProgress Then
  flex.Text = olddata
  Text1.Visible = False
End If

If Text1.Visible Then
  flex.Text = Text1.Text
End If

Text1.Visible = False
Text1.Text = ""
ChangeInProgress = False

End Sub

Private Sub Form_Activate()

If frmMain!flex.Visible = False And EditOrInsert = "Edit" Then
   MsgBox "Cannot Edit what does not exist, fool..."
   Unload frmEditInsertRow
End If

End Sub

Private Sub Form_Load()

Erase ColData
cmax = 0
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "sp_columns [" + TableName + "]"
cmd1.CommandText = Sql
Set rs = cmd1.Execute
Do While Not rs.EOF
   cmax = cmax + 1
   ColData(cmax, 1) = rs!column_name
   ColData(cmax, 2) = rs!type_name
   ColData(cmax, 3) = Str(rs!length)
    rs.MoveNext
Loop
conn.Close
If frmMain!flex.Visible = False And EditOrInsert = "Insert" Then
   flex.Rows = 2
   flex.Cols = cmax + 1
   flex.Row = 0
   For z = 1 To cmax
      flex.Col = z
      flex.Text = ColData(z, 1)
   Next
   flex.ColWidth(0) = 0
   GoTo Finish
End If

flex.Redraw = False
flex.Visible = False

flex.Rows = 2
flex.Cols = frmMain!flex.Cols
flex.Row = 0
frmMain!flex.Row = 0
For z = 1 To flex.Cols - 1
   flex.Col = z
   frmMain!flex.Col = z
   flex.Text = frmMain!flex.Text
Next

If EditOrInsert = "Edit" Then
   flex.Row = 1
   frmMain!flex.Row = CurrRow
   For z = 1 To flex.Cols - 1
      flex.Col = z
      frmMain!flex.Col = z
      flex.Text = frmMain!flex.Text
      ColData(z, 4) = flex.Text
   Next
End If

Done:
For z = 0 To flex.Cols - 1
   flex.ColWidth(z) = frmMain!flex.ColWidth(z)
Next

Finish:
flex.Redraw = True
flex.Visible = True
 Select Case EditOrInsert
   Case "Edit"
      frmEditInsertRow.Caption = "Edit Row"
      Command1.Caption = "Update"
   Case "Insert"
      frmEditInsertRow.Caption = "Insert Row"
      Command1.Caption = "Insert"
End Select

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
   Case 13
      flex.Text = Text1.Text
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
   Case 27
      flex.Text = olddata
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
End Select

End Sub

Public Sub UpdateRow()

USql = "update " + TableName + " set "
For z = 1 To flex.Cols - 1
   flex.Row = 0
   flex.Col = z
   cn$ = LCase(flex.Text)
   For zz = 1 To cmax
     If LCase(ColData(zz, 1)) = cn$ Then
        USql = USql + ColData(zz, 1) + "="
        flex.Row = 1
        flex.Col = z
        cd$ = Replace(flex.Text, "'", "''")
        Select Case ColData(zz, 2)
           Case "char", "datetime", "nchar", "image", "ntext", "nvarchar", "smalldatetime", "text", "varchar"
              USql = USql + "'" + cd$ + "',"
           Case "varbinary", "binary"
              If InStr(Trim(Str(Len(cd$)) / 2), ".") Then cd$ = "0" + cd$
              If Left(cd$, 2) <> "0x" Then cd$ = "0x" + cd$
              USql = USql + cd$ + ","
           Case "bit"
              If LCase(cd$) = "false" Then cd$ = "0"
              If LCase(cd$) = "true" Then cd$ = "1"
              USql = USql + cd$ + ","
           Case Else
              If cd$ = "" Then cd$ = "0"
              USql = USql + cd$ + ","
        End Select
     End If
   Next
Next
USql = Left(USql, Len(USql) - 1) + " where "
For z = 1 To cmax
  USql = USql + ColData(z, 1) + "="
  cd$ = Replace(ColData(z, 4), "'", "''")
  Select Case ColData(z, 2)
     Case "char", "datetime", "nchar", "image", "ntext", "nvarchar", "smalldatetime", "text", "varchar"
        USql = USql + "'" + cd$ + "'"
     Case "varbinary", "binary"
        If InStr(Trim(Str(Len(cd$)) / 2), ".") Then cd$ = "0" + cd$
        If Left(cd$, 2) <> "0x" Then cd$ = "0x" + cd$
        USql = USql + cd$
     Case "bit"
        If LCase(cd$) = "true" Then cd$ = "1"
        If LCase(cd$) = "false" Then cd$ = "0"
        USql = USql + cd$
     Case Else
        If cd$ = "" Then cd$ = "0"
        USql = USql + cd$
  End Select
  USql = USql + " and "
Next
USql = Left(USql, Len(USql) - 5)
On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = USql
cmd1.Execute
conn.Close
If Err.Number = 0 Then
   On Error GoTo 0
   Unload frmEditInsertRow
End If
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next


End Sub
Public Sub InsertRow()

ISql = "Insert into " + TableName + " ("
flex.Row = 0
For z = 1 To flex.Cols - 1
   flex.Col = z
   If InStr(flex.Text, " ") Then
      ISql = ISql + "[" + flex.Text + "],"
   Else
      ISql = ISql + flex.Text + ","
   End If
Next
ISql = Left(ISql, Len(ISql) - 1) + ") values ("

For z = 1 To flex.Cols - 1
   flex.Row = 0
   flex.Col = z
   cn$ = flex.Text
   For zz = 1 To cmax
     If ColData(zz, 1) = cn$ Then
        flex.Row = 1
        flex.Col = z
        cd$ = Replace(flex.Text, "'", "''")
        Select Case ColData(zz, 2)
           Case "char", "datetime", "nchar", "image", "ntext", "nvarchar", "smalldatetime", "text", "varchar"
              ISql = ISql + "'" + cd$ + "',"
           Case "varbinary", "binary"
              If InStr(Trim(Str(Len(cd$)) / 2), ".") Then cd$ = "0" + cd$
              If Left(cd$, 2) <> "0x" Then cd$ = "0x" + cd$
              ISql = ISql + cd$ + ","
           Case "uniqueidentifier"
              ISql = ISql + "'" + cd$ + "',"
           Case "bit"
              If LCase(cd$) = "true" Then cd$ = "1"
              If LCase(cd$) = "false" Then cd$ = "0"
              ISql = ISql + cd$ + ","
           Case Else
              If cd$ = "" Then cd$ = "0"
              ISql = ISql + cd$ + ","
        End Select
     End If
   Next
Next
ISql = Left(ISql, Len(ISql) - 1) + ")"

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = ISql
cmd1.Execute
If Err.Number = 0 Then
   conn.Close
   On Error GoTo 0
   Unload frmEditInsertRow
End If
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next

End Sub


