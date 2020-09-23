VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDebug 
   Caption         =   "Debug Stored Procedure"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14235
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Editor"
      Height          =   10455
      Left            =   3360
      TabIndex        =   10
      Top             =   0
      Width           =   10815
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   9720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Apply To Debugger"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   9720
         Width           =   1695
      End
      Begin VB.TextBox txtEd 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   9375
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Text            =   "frmDebug.frx":030A
         Top             =   240
         Width           =   10215
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   5400
   End
   Begin MSFlexGridLib.MSFlexGrid sflex 
      Height          =   3015
      Left            =   7320
      TabIndex        =   15
      Top             =   5760
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Apply to SP"
      Height          =   375
      Left            =   12480
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid rflex 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   9000
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   2566
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   13560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid vflex 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   5295
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9340
      _Version        =   393216
      GridColor       =   12582912
      GridLinesFixed  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   12720
      TabIndex        =   0
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label lbPrint 
      Height          =   255
      Left            =   7320
      TabIndex        =   17
      Top             =   8760
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "System Variables"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lbRS 
      Caption         =   "Recordset"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "F8/N-Next  F5/R-Restart  Esc/X-Quit"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Variables"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   3255
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Statements(5000) As String
Dim smax As Long
Dim TempS(5000) As String
Dim tmax As Integer
Dim Vars(500, 3) As String
Dim vmax As Integer
Dim Labels(500, 2) As String
Dim lmax As Integer
Dim Cursor(500, 3) As String
Dim cmax As Integer

Dim olddata As String
Dim ChangeInProgress As Boolean
Dim WhileInProgress As Boolean
Dim FetchStatus As Boolean
Dim ColsFetch(50) As String
Dim CurrPos As Integer

Dim dconn As ADODB.Connection
Dim dcmd1 As ADODB.Command
Dim drs As ADODB.Recordset
Dim dr1 As ADODB.Recordset
Dim dr2 As ADODB.Recordset
Dim dr3 As ADODB.Recordset
Dim dr4 As ADODB.Recordset
Dim dr5 As ADODB.Recordset

Dim EqStr As String
Dim recordcount As Integer
Dim IFStatus(500) As Integer
Dim IFInProgress As Integer
Dim WhileStartPos As Integer
Dim matheq As String
Dim FunctionPresent As Boolean
Dim OldFlexText As String
Dim StatementDone As Boolean
Dim ConvertData(2) As String


Private Sub Command1_Click()

Unload frmDebug

End Sub

Private Sub Command2_Click()

frmDebug.MousePointer = 11
Erase Statements
Erase Vars
Erase TempS
Erase Cursor
smax = 0
vmax = 0
tmax = 0
cmax = 0
flex.Clear
vflex.Clear
rflex.Clear
lbPrint.Caption = ""
WhileInProgress = False

alltext$ = ""
s$ = ""
For z = 1 To Len(txtEd.Text)
   Select Case Mid(txtEd.Text, z, 1)
      Case Chr(13)
         smax = smax + 1
          If InStr(LCase(s$), "select ") Or InStr(LCase(s$), "delete ") Or InStr(LCase(s$), "update ") Or InStr(LCase(s$), "insert ") Then SelectFlag = True
         s$ = Replace(s$, Chr$(9), "    ")
         s$ = Replace(s$, Chr$(10), "")
         s$ = Replace(s$, Chr$(13), "")
         If s$ <> "" Then
            If SelectFlag Then
               Statements(smax) = Statements(smax) + " " + s$
               smax = smax - 1
            Else
               Statements(smax) = s$
            End If
         Else
            SelectFlag = False
         End If
         alltext$ = alltext$ + s$ + " "
         s$ = ""
      Case Chr(10)
      Case Else
        s$ = s$ + Mid(txtEd.Text, z, 1)
  End Select
Next
smax = smax + 1

flex.Rows = smax + 1
flex.ColWidth(0) = 450
flex.ColWidth(1) = 25000
flex.Col = 1
startpos = 0
For z = 1 To smax
   flex.Row = z
   flex.Col = 0
   flex.Text = Str(z)
   flex.Col = 1
   flex.Text = Trim(Statements(z))
   flex.CellBackColor = "&H00C00000"
   flex.CellForeColor = "&H00FFFFFF"
   If InStr(flex.Text, " AS") Or InStr(flex.Text, " AS ") Or LCase(flex.Text) = "as" Then startpos = z + 1
Next
If startpos > smax Then
   frmDebug.MousePointer = 1
   strText = ""
   For z = 1 To smax
     strText = strText + Statements(z) + vbCrLf
     If InStr(LCase(Statements(z)), "select ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "insert ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "update ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "delete ") Then strText = strText + vbCrLf
   Next
   txtEd.Text = strText
   Frame1.Visible = True
   txtEd.SetFocus
   Exit Sub
End If

SLoop:
If startpos >= 8 And startpos + 7 <= flex.Rows - 1 Then flex.TopRow = startpos - 7
flex.Row = startpos
CurrPos = startpos
flex.Col = 0
flex.Text = "==>"
flex.Col = 1
flex.CellForeColor = "&H0000FFFF"
If flex.Text = "" Or Left(Trim(flex.Text), 2) = "/*" Or Left(Trim(flex.Text), 2) = "--" Or Right(Trim(flex.Text), 1) = ":" Then
   flex.Col = 0
   flex.Text = ""
   flex.CellForeColor = "&H00FFFFFF"
   startpos = startpos + 1
   GoTo SLoop
End If
Timer1_Timer

'get variables
pos = 0
VLoop:
pos = InStr(pos + 1, alltext$, "@")
If pos Then
   vn$ = ""
   p = pos
   Do Until Mid(alltext$, p, 1) = " " Or InStr(EqStr, Mid(alltext$, p, 1)) Or Mid(alltext$, p, 1) = "(" Or Mid(alltext$, p, 1) = ")" Or Mid(alltext$, p, 1) = "," Or Mid(alltext$, p, 1) = Chr(13) Or p > Len(alltext$)
      vn$ = vn$ + Mid(alltext$, p, 1)
      p = p + 1
   Loop
   If InStr(vn$, " ") Then vn$ = Left$(vn$, InStr(vn$, " ") - 1)
   p = p + 1
   If Left(vn$, 2) = "@@" Then
      pos = pos + 1
      GoTo VLoop
   End If
   'get type
GetTypeLoop:
   vt$ = ""
   Do Until Mid(alltext$, p, 1) = " " Or InStr(EqStr, Mid(alltext$, p, 1)) Or Mid(alltext$, p, 1) = "," Or Mid(alltext$, p, 1) = Chr(13) Or p > Len(alltext$)
      vt$ = vt$ + Mid(alltext$, p, 1)
      p = p + 1
   Loop
   vt$ = LCase(vt$)
   If vt$ = "as" Then
      p = p + 1
      GoTo GetTypeLoop
   End If
   If InStr(vt$, "(") Then
      vt$ = Left(vt$, InStr(vt$, "(") - 1)
   End If
   Select Case vt$
      Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
         defaultvalue$ = "0"
      Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
         defaultvalue$ = "<NULL>"
      Case "binary", "image"
         defaultvalue$ = "x0"
   End Select
   
   flag% = 0
   For z = 1 To vmax
      If vn$ = Vars(z, 1) Then
         flag% = 1
         Exit For
      End If
   Next
   If flag% = 0 Then
       vmax = vmax + 1
       Vars(vmax, 1) = vn$
       Vars(vmax, 2) = vt$
       Vars(vmax, 3) = defaultvalue$
   End If
   GoTo VLoop
End If

Call UpdateVariables
frmDebug.MousePointer = 1
Frame1.Visible = False
flex.SetFocus
flex.TopRow = 1

End Sub

Private Sub Command3_Click()

Frame1.Visible = False
flex.SetFocus

End Sub

Private Sub Command4_Click()

msg$ = "The original format such as mutiline selects, updates etc, will be lost applying at this time.  Are you sure you want to apply the changes now?"
sysbut% = MsgBox(msg$, 4)
If sysbut% <> vbYes Then Exit Sub

On Error GoTo errorfound
strText = ""
flex.Col = 1
For z = 1 To flex.Rows - 1
   flex.Row = z
   strText = strText + flex.Text + vbCrLf
   If InStr(LCase(Statements(z)), "select ") Then strText = strText + vbCrLf
   If InStr(LCase(Statements(z)), "insert ") Then strText = strText + vbCrLf
   If InStr(LCase(Statements(z)), "update ") Then strText = strText + vbCrLf
   If InStr(LCase(Statements(z)), "delete ") Then strText = strText + vbCrLf
Next
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = strText
cmd1.Execute
On Error GoTo 0
flex.SetFocus
Exit Sub

errorfound:
   MsgBox "Unable to Apply due to - " + Err.Description
   Resume Next

End Sub

Private Sub flex_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
   Case 119 'next
       If CurrPos <= flex.Rows - 1 Then
          flex.Row = CurrPos
          flex.Col = 0
          flex.Text = ""
          flex.Col = 1
          flex.CellForeColor = "&H00FFFFFF"
          Call ExecuteStatement
          flex.Row = CurrPos
          flex.Col = 0
          flex.Text = "==>"
          flex.Col = 1
          flex.CellForeColor = "&H0000FFFF"
          flex.Refresh
SLoop:
       If CurrPos >= 8 And CurrPos + 7 <= flex.Rows - 1 Then flex.TopRow = CurrPos - 7
       flex.Row = CurrPos
       flex.Col = 0
       flex.Text = "==>"
       flex.Col = 1
       flex.CellForeColor = "&H0000FFFF"
       If flex.Text = "" Or Left(Trim(flex.Text), 2) = "/*" Or Left(Trim(flex.Text), 2) = "--" Or Right(Trim(flex.Text), 1) = ":" Then
          flex.Col = 0
          flex.Text = ""
          flex.CellForeColor = "&H00FFFFFF"
          CurrPos = CurrPos + 1
          If CurrPos > flex.Rows - 1 Then GoTo Done
          GoTo SLoop
       End If
       End If
   Case 116 'restart
      List1_Click
End Select

Done:
flex.Refresh
flex.SetFocus

End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
   Case 110, 78 'next
       If CurrPos <= flex.Rows - 1 Then
       flex.Row = CurrPos
       flex.Col = 0
       flex.Text = ""
       flex.Col = 1
       flex.CellForeColor = "&H00FFFFFF"
       Call ExecuteStatement
       flex.Row = CurrPos
       flex.Col = 0
       flex.Text = "==>"
       flex.Col = 1
       flex.CellForeColor = "&H0000FFFF"
       flex.Refresh
SLoop:
       If CurrPos >= 8 And CurrPos + 7 <= flex.Rows - 1 Then flex.TopRow = CurrPos - 7
       flex.Row = CurrPos
       flex.Col = 0
       flex.Text = "==>"
       flex.Col = 1
       flex.CellForeColor = "&H0000FFFF"
       If flex.Text = "" Or Left(Trim(flex.Text), 2) = "/*" Or Left(Trim(flex.Text), 2) = "--" Or Right(Trim(flex.Text), 1) = ":" Then
          flex.Col = 0
           flex.Text = ""
          flex.CellForeColor = "&H00FFFFFF"
          CurrPos = CurrPos + 1
          If CurrPos > flex.Rows - 1 Then GoTo Done
          GoTo SLoop
       End If
       End If
   Case 115, 83 'stop
   Case 114, 82 'restart
      List1_Click
   Case 27, 120, 88 'exit
      Unload frmDebug
      Exit Sub
End Select

Done:
flex.Refresh
flex.SetFocus

End Sub

Private Sub Form_Load()

Frame1.Visible = False
List1.Clear
EqStr = "=<>"
matheq = "+-*"
For z = 0 To frmMain!List4.ListCount - 1
    If frmMain!List4.List(z) <> "sp_AdminIFCheck" Then List1.AddItem frmMain!List4.List(z)
Next
rflex.ColWidth(0) = 0

Set dconn = New ADODB.Connection
Set dcmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
dconn.Open strConnect
dcmd1.ActiveConnection = dconn

sflex.Rows = 7
sflex.Cols = 2
sflex.ColWidth(0) = 1500
sflex.ColWidth(1) = 2000
sflex.Row = 0
sflex.Col = 0
sflex.Text = "Variable"
sflex.Col = 1
sflex.Text = "Value"

sflex.Col = 0
sflex.Row = 1
sflex.Text = "RecordCount"
sflex.Col = 1
sflex.Text = " 0"

sflex.Col = 0
sflex.Row = 2
sflex.Text = "@@rowcount"
sflex.Col = 1
sflex.Text = " 0"

sflex.Col = 0
sflex.Row = 3
sflex.Text = "@@identity"
sflex.Col = 1
sflex.Text = "<NULL>"

sflex.Col = 0
sflex.Row = 4
sflex.Text = "@@error"
sflex.Col = 1
sflex.Text = " 0"

sflex.Col = 0
sflex.Row = 5
sflex.Text = "ErrorDesc"
sflex.Col = 1
sflex.Text = "<NULL>"

sflex.Col = 0
sflex.Row = 6
sflex.Text = "Return"
sflex.Col = 1
sflex.Text = " 0"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

dconn.Close

End Sub

Private Sub List1_Click()

frmDebug.MousePointer = 11
For z = 0 To List1.ListCount - 1
  If List1.Selected(z) Then
     SPName = List1.List(z)
     Exit For
  End If
Next

Erase Statements
Erase Vars
Erase TempS
Erase IFStatus
Erase Cursor
smax = 0
vmax = 0
tmax = 0
cmax = 0
IFSPMax = 0
flex.Clear
vflex.Clear
rflex.Clear
lbRS.Caption = "Recordset"
lbPrint.Caption = ""
WhileInProgress = False

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = "sp_helptext " + SPName
Set rs = cmd1.Execute
alltext$ = ""
s$ = ""
Do While Not rs.EOF
    smax = smax + 1
    s$ = RTrim(rs!Text)
    If Right(s$, 3) = ":" + vbCrLf Or Right(s$, 1) = ":" Then
        s$ = Trim(s$)
        s$ = Left(s$, Len(s$) - 1)
        lmax = lmax + 1
        Labels(lmax, 1) = s$
        Labels(lmax, 2) = Str(smax)
    End If
    s$ = Replace(s$, "CREATE PROCEDURE", "ALTER PROCEDURE")
    If InStr(LCase(s$), "select ") Or InStr(LCase(s$), "delete ") Or InStr(LCase(s$), "update ") Or InStr(LCase(s$), "insert ") Or (InStr(LCase(s$), "declare ") And InStr(LCase(s$), " cursor ")) Then SelectFlag = True
    s$ = Replace(s$, Chr$(9), "    ")
    s$ = Replace(s$, Chr$(10), "")
    s$ = Replace(s$, Chr$(13), "")
    If s$ <> "" Then
       If SelectFlag Then
          Statements(smax) = Statements(smax) + " " + s$
          smax = smax - 1
       Else
          Statements(smax) = s$
       End If
    Else
       SelectFlag = False
    End If
    alltext$ = alltext$ + rs!Text
    rs.MoveNext
Loop
smax = smax + 1

sflex.Col = 1
sflex.Row = 1
sflex.Text = " 0"
sflex.Row = 2
sflex.Text = " 0"
sflex.Row = 3
sflex.Text = "<NULL>"
sflex.Row = 4
sflex.Text = " 0"
sflex.Row = 5
sflex.Text = "<NULL>"
sflex.Row = 6
sflex.Text = " 0"

flex.Rows = smax + 1
flex.ColWidth(0) = 450
flex.ColWidth(1) = 25000
flex.Col = 1
startpos = 0
For z = 1 To smax
   flex.Row = z
   flex.Col = 0
   flex.Text = Str(z)
   flex.Col = 1
   flex.Text = Trim(Statements(z))
   flex.CellBackColor = "&H00C00000"
   flex.CellForeColor = "&H00FFFFFF"
   If InStr(flex.Text, " AS") Or InStr(flex.Text, " AS ") Or LCase(flex.Text) = "as" Then startpos = z + 1
Next
If startpos >= smax Then
   frmDebug.MousePointer = 1
   strText = ""
   For z = 1 To smax
     strText = strText + Statements(z) + vbCrLf
     If InStr(LCase(Statements(z)), "select ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "insert ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "update ") Then strText = strText + vbCrLf
     If InStr(LCase(Statements(z)), "delete ") Then strText = strText + vbCrLf
   Next
   txtEd.Text = strText
   Frame1.Visible = True
   txtEd.SetFocus
   Exit Sub
End If

SLoop:
If startpos >= 8 And startpos + 7 <= flex.Rows - 1 Then flex.TopRow = startpos - 7
flex.Row = startpos
CurrPos = startpos
flex.Col = 0
flex.Text = "==>"
flex.Col = 1
flex.CellForeColor = "&H0000FFFF"
If flex.Text = "" Or Left(Trim(flex.Text), 2) = "/*" Or Left(Trim(flex.Text), 2) = "--" Or Right(Trim(flex.Text), 1) = ":" Then
   flex.Col = 0
   flex.Text = ""
   flex.CellForeColor = "&H00FFFFFF"
   startpos = startpos + 1
   GoTo SLoop
End If
Timer1_Timer

'get variables
pos = 0
VLoop:
pos = InStr(pos + 1, alltext$, "@")
If pos Then
   vn$ = ""
   p = pos
   Do Until Mid(alltext$, p, 1) = " " Or InStr(EqStr, Mid(alltext$, p, 1)) Or Mid(alltext$, p, 1) = "(" Or Mid(alltext$, p, 1) = ")" Or Mid(alltext$, p, 1) = "," Or Mid(alltext$, p, 1) = Chr(13) Or p > Len(alltext$)
      vn$ = vn$ + Mid(alltext$, p, 1)
      p = p + 1
   Loop
   If InStr(vn$, " ") Then vn$ = Left$(vn$, InStr(vn$, " ") - 1)
   p = p + 1
   If Left(vn$, 2) = "@@" Then
      pos = pos + 1
      GoTo VLoop
   End If
   'get type
GetTypeLoop:
   vt$ = ""
   Do Until Mid(alltext$, p, 1) = " " Or InStr(EqStr, Mid(alltext$, p, 1)) Or Mid(alltext$, p, 1) = "," Or Mid(alltext$, p, 1) = Chr(13) Or p > Len(alltext$)
      vt$ = vt$ + Mid(alltext$, p, 1)
      p = p + 1
   Loop
   vt$ = LCase(vt$)
   If vt$ = "as" Then
      p = p + 1
      GoTo GetTypeLoop
   End If
   If InStr(vt$, "(") Then
      vt$ = Left(vt$, InStr(vt$, "(") - 1)
   End If
   Select Case vt$
      Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
         defaultvalue$ = "0"
      Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
         defaultvalue$ = "<NULL>"
      Case "binary", "image"
         defaultvalue$ = "x0"
   End Select
   
   flag% = 0
   For z = 1 To vmax
      If vn$ = Vars(z, 1) Then
         flag% = 1
         Exit For
      End If
   Next
   If flag% = 0 Then
       vmax = vmax + 1
       Vars(vmax, 1) = vn$
       Vars(vmax, 2) = vt$
       Vars(vmax, 3) = defaultvalue$
   End If
   GoTo VLoop
End If

Call UpdateVariables
frmDebug.MousePointer = 1
flex.SetFocus
flex.TopRow = 1

End Sub

Private Sub rflex_DblClick()

frmDebug.MousePointer = 11
frmView.flex2.Clear
frmView.flex2.Rows = rflex.Rows
frmView.flex2.Cols = rflex.Cols
For z = 0 To rflex.Cols - 1
   frmView.flex2.ColWidth(z) = rflex.ColWidth(z)
Next

For z = 0 To rflex.Rows - 1
  rflex.Row = z
  For zz = 0 To rflex.Cols - 1
     rflex.Col = zz
     frmView.flex2.Row = z
     frmView.flex2.Col = zz
     frmView.flex2.Text = rflex.Text
  Next
Next

frmDebug.MousePointer = 1
frmView.Show 1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
   Case 13
      vflex.Text = Text1.Text
      Vars(CurrRow, 3) = Text1.Text
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
   Case 27
      vflex.Text = olddata
      Vars(CurrRow, 3) = olddata
      Text1.Visible = False
      ChangeInProgress = False
      Text1.Text = ""
End Select
If ChangeInProgress = False Then flex.SetFocus

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
   Case 13
      flex.Text = Text2.Text
      Statements(CurrRow) = Text2.Text
      Text2.Visible = False
      ChangeInProgress = False
      Text2.Text = ""
   Case 27
      flex.Text = olddata
      Statements(CurrRow) = olddata
      Text2.Visible = False
      ChangeInProgress = False
      Text2.Text = ""
End Select
flex.SetFocus

End Sub

Private Sub Timer1_Timer()

cr = flex.Row
cl = flex.Col

flex.Redraw = False
For z = 1 To flex.Rows - 1
  flex.Row = z
  If z <> CurrPos Then
     flex.Col = 0
     flex.Text = Str(z)
     flex.CellForeColor = "&H80000012"
     flex.Col = 1
     If flex.CellForeColor <> "&H00FFFFFF" Then
       flex.CellForeColor = "&H00FFFFFF"
     End If
  End If
Next
flex.Redraw = True

flex.Row = cr
flex.Col = cl

End Sub

Private Sub vflex_Click()

If vflex.MouseRow = 0 Then Exit Sub
olddata = vflex.Text
Text1.Text = ""
Text1.SelStart = 0
Text1.Height = vflex.CellHeight
Text1.Width = vflex.CellWidth
Text1.Move vflex.CellLeft + vflex.Left, vflex.CellTop + vflex.Top, vflex.CellWidth, vflex.CellHeight
Text1.Visible = True
Text1.SetFocus
ChangeInProgress = True

End Sub
Private Sub flex_Click()

strText = ""
For z = 1 To smax
  strText = strText + Trim(Statements(z)) + vbCrLf
  If InStr(LCase(Statements(z)), "select ") Then strText = strText + vbCrLf
  If InStr(LCase(Statements(z)), "insert ") Then strText = strText + vbCrLf
  If InStr(LCase(Statements(z)), "update ") Then strText = strText + vbCrLf
  If InStr(LCase(Statements(z)), "delete ") Then strText = strText + vbCrLf
Next
txtEd.Text = strText
Frame1.Visible = True
txtEd.SetFocus

End Sub

Private Sub vflex_LeaveCell()

On Error Resume Next
If Text1.Text = "" And ChangeInProgress Then
  vflex.Text = olddata
  Text1.Visible = False
End If

If Text1.Visible Then
  vflex.Text = Text1.Text
End If

Text1.Visible = False
Text1.Text = ""
ChangeInProgress = False
flex.SetFocus
On Error GoTo 0

End Sub
Private Sub vflex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

CurrRow = vflex.MouseRow
CurrCol = 1
If CurrRow = 0 Then Exit Sub

End Sub
Public Sub ExecuteStatement()

recordcount = 0
StatementDone = False
If CurrPos = flex.Rows Then Exit Sub

flex.Col = 1
cs$ = LCase(Trim(flex.Text))
If cs$ = "" Then GoTo Final

'statements start
Call CheckForFunctions(cs$)

flag% = 0
For z = 1 To 3
   If InStr(cs$, Mid(matheq, z, 1)) Then
      If InStr(cs$, "select ") = 0 And Mid(matheq, z, 1) <> "*" Then
         flag% = 1
      End If
   End If
Next
If flag% And LastFunction <> "convert" Then Call DoEval(cs$)

Cont1:
If StatementDone Then
   StatementDone = False
   GoTo Final
End If

If cs$ = "begin" Or cs$ = "go" Then GoTo Final

If Left$(cs$, 8) = "declare " Then
   If InStr(cs$, " cursor ") Then
      cmax = cmax + 1
      vn$ = ""
      cs$ = Trim(flex.Text)
      cs$ = Trim(Mid(cs$, 9))
      p = 1
      Do Until Mid(cs$, p, 1) = " "
         vn$ = vn$ + Mid(cs$, p, 1)
         p = p + 1
      Loop
      Cursor(cmax, 1) = vn$
      pos = InStr(LCase(cs$), "for ")
      cs$ = Trim(Mid(cs$, pos + 3))
      Cursor(cmax, 2) = cs$
      GoTo Final
   Else
       GoTo Done
   End If
   GoTo Done
End If

If cs$ = "end" Then
  Call DoEnd
  GoTo Final
End If

If Left$(cs$, 5) = "exec " Or Left$(cs$, 8) = "execute " Then
   Call DoExec
   GoTo Final
End If

If Left$(cs$, 6) = "fetch " Then
   Call DoFetchCursor
   GoTo Final
End If

If Left(cs$, 5) = "goto " Then
  cs$ = flex.Text
  For z = 1 To lmax
     If Left(Labels(z, 1), InStr(Labels(z, 1), ":") - 1) = Trim(Mid(cs$, 5)) Then
         flex.Row = CurrPos
         flex.Col = 0
         flex.Text = ""
         flex.Col = 1
         flex.CellForeColor = "&H00FFFFFF"
         CurrPos = Val(Labels(z, 2))
         flex.Row = CurrPos
         flex.Col = 0
         flex.Text = "==>"
         flex.Col = 1
         flex.CellForeColor = "&H0000FFFF"
         flex.Refresh
         Exit For
      End If
    Next
    GoTo Final
End If

If Left$(cs$, 3) = "if " Then
   Call DoIFStatement(CurrPos)
   GoTo Final
End If

If Left(cs$, 12) = "insert into " Or Left(cs$, 7) = "update " Or Left(cs$, 7) = "delete " Then
   Call DoInsertUpdateDelete(cs$)
   GoTo Final
End If

If Left$(cs$, 5) = "open " Then
   Call DoOpenCursor
   GoTo Final
End If

If Left$(cs$, 6) = "print " Then
   Call DoPrint
   GoTo Final
End If

If Left$(cs$, 7) = "return " Then
   Call DoReturn
   GoTo Final
End If

If Left(cs$, 7) = "select " Then
   Call DoSelect(cs$)
   GoTo Final
End If

If Left$(cs$, 4) = "set " Then
   Call DoSet(cs$)
   GoTo Final
End If

If Left(cs$, 4) = "use " Then
   NewDBName = Trim(Mid(cs$, 5))
   dconn.Close
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + NewDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   On Error GoTo errorfound
   dconn.Open strConnect
   dcmd1.ActiveConnection = dconn
   On Error GoTo 0
   GoTo Final
End If

If Left$(cs$, 6) = "while " Then
   Call DoWhile
   GoTo Final
End If

If Left$(cs$, 6) = "break " And WhileInProgress Then
   flex.Col = 1
   For z = 1 To flex.Rows - 1
      flex.Row = z
      cs$ = Trim(LCase(flex.Text))
      If cs$ = "end" Then
         CurrPos = z
         Exit For
      End If
   Next
   WhileInProgress = False
   GoTo Final
End If

If Left$(cs$, 9) = "continue " And WhileInProgress Then
   CurrPos = WhileStartPos - 1
   GoTo Final
End If

Done:
dcmd1.CommandText = flex.Text
flex.ToolTipText = flex.Text
dcmd1.Execute

Final:
If FunctionPresent Then
   FunctionPresent = False
   flex.Text = OldFlexText
End If
CurrPos = CurrPos + 1
If CurrPos = flex.Rows Then CurrPos = CurrPos - 1
On Error GoTo 0
Timer1_Timer
Exit Sub

errorfound:
   If Err.Description <> olderr Then
        MsgBox Err.Description
        olderr = Err.Description
   End If
   Resume Next
   
End Sub

Public Sub DoSelect(cs$)

Dim p1 As Integer
Dim p2 As Integer
Dim p3 As Integer

   rflex.Clear
   'check for set @ = format
   p1 = InStr(cs$, "@")
   p2 = InStr(cs$, "=")
   If p2 = 0 Then p2 = InStr(cs$, "<")
   If p2 = 0 Then p2 = InStr(cs$, ">")
   If p1 > 0 And p2 > 0 Then
      If p1 < p2 Then
         Call DoSelectEqual(p1, p2)
         Exit Sub
      End If
   End If
   'check for set @ as format
   p1 = InStr(cs$, "@")
   p2 = InStr(cs$, "=")
   If p2 = 0 Then p2 = InStr(cs$, "<")
   If p2 = 0 Then p2 = InStr(cs$, ">")
   p3 = InStr(cs$, " as ")
   If p1 And p3 And p2 = 0 Then
      'Call DoSelectReturn
      Exit Sub
   End If

RegularSelect:
   
   For z = 1 To vmax
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, LCase(Vars(z, 1)), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, LCase(Vars(z, 1)), "is null")
            Else
                cs$ = Replace(cs$, LCase(Vars(z, 1)), "'" + Vars(z, 3) + "'")
            End If
         'Case "binary", "image"
      End Select
   Next
   dcmd1.CommandText = cs$ + vbCrLf + vbCrLf + "Select @@rowcount as sv1, @@identity as sv2, @@error as sv3"
   flex.ToolTipText = cs$
   Set drs = dcmd1.Execute
   rflex.Cols = drs.Fields.Count + 1
   rflex.Row = 0
   cr& = 0
   smax2 = 0
   For z = 0 To drs.Fields.Count - 1
      rflex.Col = z + 1
      rflex.Text = LCase(drs.Fields(z).Name)
   Next
   'do data
   Do While Not drs.EOF
      cr& = cr& + 1
      rflex.Rows = cr& + 1
      rflex.Row = cr&
      For z = 0 To drs.Fields.Count - 1
         rflex.Col = z + 1
         Select Case drs.Fields(z).Type
            Case 3, 5, 131, 2, 6, 17, 4, 20
               If IsNull(drs.Fields(z).Value) Then
                  rflex.Text = " 0"
               Else
                  rflex.Text = " " + Trim(Str(drs.Fields(z).Value))
               End If
            Case 204, 128
               acnv = StrConv(drs.Fields(z).Value, vbUnicode)
               If IsNull(acnv) Then
                  hxstr = ""
               Else
                  hxstr = " 0x"
                  For y = 1 To Len(acnv)
                     hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
                  Next
               End If
               rflex.Text = " " + Trim(hxstr)
            Case 129, 200, 135, 202, 203, 11, 72, 201, 205, 130
               If IsNull(drs.Fields(z).Value) Then
                  rflex.Text = ""
                  If NullReq = "Y" Then
                     rflex.Text = "Null"
                  End If
               Else
                  rflex.Text = drs.Fields(z).Value
                  If IsNumeric(Left(drs.Fields(z).Value, 1)) Then rflex.Text = " " + rflex.Text
               End If
            Case Else
               MsgBox "Unable to resolve type " + Str(drs.Fields(z).Type) + " on column " + drs.Fields(z).Name
         End Select
      Next
      drs.MoveNext
   Loop
   recordcount = cr&
   lbRS.Caption = "Recordset (" + Trim(Str(cr&)) + " rows)"
   Set drs = drs.NextRecordset
Call UpdateSystemVariables

End Sub

Public Sub DoSet(cs$)

cs$ = flex.Text
If InStr(cs$, "@") = 0 Then Exit Sub
cs$ = Trim(Mid(cs$, InStr(cs$, "@")))
pos = InStr(cs$, "=")
If pos = 0 Then Exit Sub

csL$ = Trim(Left(cs$, pos - 1))
csR$ = Trim(Mid(cs$, pos + 1))

If Left(csR$, 1) = "@" And Mid(csR$, 2, 1) <> "@" Then
   For z = 1 To vmax
      If Vars(z, 1) = csL$ Then
         pos = z
         Exit For
      End If
   Next
   For z = 1 To vmax
      If Vars(z, 1) = csR$ Then
         Vars(pos, 3) = Vars(z, 3)
         Exit For
      End If
   Next
Else
   For z = 1 To vmax
      If Vars(z, 1) = csL$ Then
         If Left(csR$, 1) = "'" And Right(csR$, 1) = "'" Then
            Vars(z, 3) = Replace(csR$, "'", "")
         Else
            Vars(z, 3) = csR$
         End If
         Exit For
      End If
   Next
End If
Call UpdateVariables

End Sub

Public Sub UpdateVariables()

vflex.Rows = vmax + 1
vflex.Cols = 2
vflex.ColWidth(0) = 2000
vflex.ColWidth(1) = 5000
vflex.Row = 0
vflex.Col = 0
vflex.Text = "Name"
vflex.Col = 1
vflex.Text = "Value"
For z = 1 To vmax
    vflex.Row = z
    vflex.Col = 0
    vflex.Text = Vars(z, 1)
    vflex.Col = 1
    If Vars(z, 3) = "" Then
       vflex.Text = "<NULL>"
    Else
       vflex.Text = Vars(z, 3)
       If IsNumeric(Left(vflex.Text, 1)) Then vflex.Text = " " + Vars(z, 3)
    End If
Next

End Sub

Public Sub DoInsertUpdateDelete(cs$)

On Error GoTo errorfound
If FunctionPresent = False Then
   cs$ = flex.Text
   For z = 1 To vmax
     If InStr(cs$, Vars(z, 1)) Then
         Select Case Vars(z, 2)
            Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
               cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
            Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
               If Vars(z, 3) = "<NULL>" Then
                   cs$ = Replace(cs$, Vars(z, 1), "''")
               Else
                   cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
               End If
            'Case "binary", "image"
         End Select
     End If
   Next
End If

dcmd1.CommandText = cs$ + vbCrLf + vbCrLf + "Select @@rowcount as sv1, @@identity as sv2, @@error as sv3"
flex.ToolTipText = cs$
Set drs = dcmd1.Execute
Set drs = drs.NextRecordset
On Error GoTo 0
Call UpdateSystemVariables
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next

End Sub

Public Sub DoSelectEqual(p1 As Integer, p2 As Integer)

On Error GoTo errorfound

If FunctionPresent = False Then cs$ = flex.Text
p1 = InStr(cs$, "@")
p2 = InStr(cs$, "=")
currvar = Trim(Mid(cs$, p1, p2 - p1))
cs$ = Trim(Mid$(cs$, p2 + 1))
If Left(cs$, 1) = "=" Then cs$ = Right(cs$, Len(cs$) - 1)

For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
         'Case "binary", "image"
      End Select
  End If
Next
  
If Val(cs$) Or Left(cs$, 1) = "'" Then
   For z = 1 To vmax
      If Vars(z, 1) = currvar Then
         If Left(cs$, 1) = "'" And Right(cs$, 1) = "'" Then
            Vars(z, 3) = Replace(cs$, "'", "")
         Else
            Vars(z, 3) = cs$
         End If
         Exit For
      End If
   Next
   Call UpdateVariables
Else
   Sql = "Select " + cs$
   flex.ToolTipText = Sql
   dcmd1.CommandText = Sql + vbCrLf + vbCrLf + "Select @@rowcount as sv1, @@identity as sv2, @@error as sv3"
   Set drs = dcmd1.Execute
   recordcount = drs.recordcount
   For z = 1 To vmax
      If currvar = Vars(z, 1) Then
         Select Case Vars(z, 2)
            Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
               If drs.BOF Or drs.EOF Then
                  Vars(z, 3) = "<NULL>"
               Else
                  Vars(z, 3) = Str(drs.Fields(0).Value)
               End If
            Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
               If drs.BOF Or drs.EOF Then
                  Vars(z, 3) = "<NULL>"
               Else
                  If IsNull(drs.Fields(0).Value) Then
                      Vars(z, 3) = "<NULL>"
                  Else
                      Vars(z, 3) = drs.Fields(0).Value
                  End If
               End If
           'Case "binary", "image"
         End Select
      End If
   Next
   Set drs = drs.NextRecordset
   Call UpdateVariables
   Call UpdateSystemVariables
End If
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next
   
End Sub

Public Sub UpdateSystemVariables()

sflex.Row = 1
sflex.Col = 1
sflex.Text = " " + Str(recordcount)

sflex.Row = 2
sflex.Col = 1
sflex.Text = Str(drs!sv1)

sflex.Row = 3
sflex.Col = 1
If IsNull(drs!sv2) Then
   sflex.Text = "<NULL>"
Else
   sflex.Text = " " + drs!sv2
End If

sflex.Row = 4
sflex.Col = 1
If IsNull(drs!sv3) Then
   sflex.Text = "<NULL>"
Else
   sflex.Text = " " + Str(drs!sv3)
End If
desctext$ = "<NULL>"
If Val(sflex.Text) Then
   Sql = "Select description from master.dbo.sysmessages where error=" + sflex.Text
   dcmd1.CommandText = Sql
   Set rs = dcmd1.Execute
   If rs.BOF Or rs.EOF Then
      desctext$ = "Unable to locate Error description."
   Else
      desctext$ = rs!Description
   End If
End If
sflex.Row = 5
sflex.Col = 1
sflex.Text = desctext$

End Sub

Public Sub DoIFStatement(cp%)

On Error GoTo errorfound
If FunctionPresent = False Then cs$ = flex.Text

For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
         'Case "binary", "image"
      End Select
  End If
Next

flag% = 0
For z = 1 To IFSPMax
   If IFStoredProc(z, 1) = Server And IFStoredProc(z, 2) = DBName Then
      flag% = 1
      Exit For
   End If
Next
If flag% = 0 Then
   IFSPMax = IFSPMax + 1
   IFStoredProc(IFSPMax, 1) = Server
   IFStoredProc(IFSPMax, 2) = DBName
End If

Sql = "CREATE PROCEDURE sp_AdminIFCheck" + vbCrLf
Sql = Sql + "AS" + vbCrLf
Sql = Sql + "declare @IFStatementReturn int" + vbCrLf
Sql = Sql + cs$ + vbCrLf
Sql = Sql + "begin" + vbCrLf
Sql = Sql + "set @IFStatementReturn=1" + vbCrLf
Sql = Sql + "end" + vbCrLf
Sql = Sql + "else" + vbCrLf
Sql = Sql + "begin" + vbCrLf
Sql = Sql + "set @IFStatementReturn=0" + vbCrLf
Sql = Sql + "end" + vbCrLf
Sql = Sql + "Select @IFStatementReturn as ISF"
Redo:
dcmd1.CommandText = Sql
dcmd1.Execute
dcmd1.CommandText = "sp_AdminIFCheck"
Set drs = dcmd1.Execute
IFInProgress = IFInProgress + 1
If drs.Fields(0).Value Then
   IFStatus(IFInProgress) = 1
Else
   IFStatus(IFInProgress) = 0
   flex.Col = 1
   begin = 0
   For z = CurrPos + 1 To flex.Rows - 1
      flex.Row = z
      cs$ = Trim(LCase(flex.Text))
      If cs$ = "begin" Then begin = begin + 1
      If cs$ = "end" Then
         begin = begin - 1
         If begin = 0 Then
            CurrPos = z
            flex.Row = z + 1
            If Trim(LCase(flex.Text)) = "else" Then CurrPos = z + 1
            Exit For
         End If
      End If
ContEnd:
   Next
End If
On Error GoTo 0
Exit Sub

errorfound:
  If InStr(LCase(Err.Description), "already an object") Then
     Sql = Replace(Sql, "CREATE PROCEDURE", "ALTER PROCEDURE")
     Resume Redo
  End If
  MsgBox "The following error occurred - " + Err.Description
  Resume Next
  
End Sub

Public Sub DoReturn()

cs$ = Trim(flex.Text)
For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
         'Case "binary", "image"
      End Select
  End If
Next

rv = Mid(cs$, 7)
sflex.Row = 6
sflex.Col = 1
sflex.Text = rv
CurrPos = smax

End Sub

Public Sub DoOpenCursor()

cs$ = flex.Text
vn$ = Trim(Mid(cs$, InStr(LCase(cs$), "open ") + 5))
openedcur = 0
For z = 1 To cmax
   If Cursor(z, 3) <> "" Then
      openedcur = openedcur + 1
   End If
Next
For z = 1 To cmax
    If Cursor(z, 1) = vn$ Then
       Sql = Cursor(z, 2)
       openedcur = openedcur + 1
       Cursor(z, 3) = Str(openedcur)
       dcmd1.CommandText = Sql
       Select Case openedcur
           Case 1
              Set dr1 = dcmd1.Execute
              FetchStatus = dr1.EOF
           Case 2
              Set dr2 = dcmd1.Execute
              FetchStatus = dr2.EOF
           Case 3
              Set dr3 = dcmd1.Execute
              FetchStatus = dr3.EOF
           Case 4
              Set dr4 = dcmd1.Execute
              FetchStatus = dr4.EOF
           Case 5
              Set dr5 = dcmd1.Execute
              FetchStatus = dr5.EOF
           Case Else
              MsgBox "Only five cursors can be opened in this admin tool..."
       End Select
       Exit For
    End If
Next

End Sub

Public Sub DoFetchCursor()

On Error GoTo errorfound

cs$ = Trim(flex.Text)
p = InStr(LCase(cs$), " from ") + 5
cs$ = Trim(Mid(cs$, p))
vn$ = Left(cs$, InStr(cs$, " ") - 1)
p = InStr(LCase(cs$), " into ") + 5
scol$ = Trim(Mid(cs$, p))

For z = 1 To cmax
   If vn$ = Cursor(z, 1) Then
      cnum = z
      Exit For
   End If
Next
Erase ColsFetch
cfmax = 0

s$ = ""
For z = 1 To Len(scol$)
   Select Case Mid(scol$, z, 1)
      Case ","
         cfmax = cfmax + 1
         ColsFetch(cfmax) = s$
         s$ = ""
      Case " " 'do nothing
      Case Else
        s$ = s$ + Mid(scol$, z, 1)
   End Select
Next
cfmax = cfmax + 1
ColsFetch(cfmax) = s$
fpos = 0

Select Case cnum
   Case 1
      For z = 1 To cfmax
         Call GetVariablePositionAndType(ColsFetch(z), pos%, ty%)
         If ty% = 1 Then
            If IsNull(dr1.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = dr1.Fields(fpos).Value
            End If
         Else
            If IsNull(dr1.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = Str(dr1.Fields(fpos).Value)
            End If
         End If
         fpos = fpos + 1
      Next
      FetchStatus = dr1.EOF
      If FetchStatus = False Then dr1.MoveNext
   Case 2
      For z = 1 To cfmax
         Call GetVariablePositionAndType(ColsFetch(z), pos%, ty%)
         If ty% = 1 Then
            If IsNull(dr2.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = dr2.Fields(fpos).Value
            End If
         Else
            If IsNull(dr2.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = Str(dr2.Fields(fpos).Value)
            End If
         End If
         fpos = fpos + 1
      Next
      FetchStatus = dr2.EOF
      If FetchStatus = False Then dr2.MoveNext
   Case 3
      For z = 1 To cfmax
         Call GetVariablePositionAndType(ColsFetch(z), pos%, ty%)
         If ty% = 1 Then
            If IsNull(dr3.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = dr3.Fields(fpos).Value
            End If
         Else
            If IsNull(dr3.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = Str(dr3.Fields(fpos).Value)
            End If
         End If
         fpos = fpos + 1
      Next
      FetchStatus = dr3.EOF
      If FetchStatus = False Then dr3.MoveNext
   Case 4
      For z = 1 To cfmax
         Call GetVariablePositionAndType(ColsFetch(z), pos%, ty%)
         If ty% = 1 Then
            If IsNull(dr4.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = dr4.Fields(fpos).Value
            End If
         Else
            If IsNull(dr4.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = Str(dr4.Fields(fpos).Value)
            End If
         End If
         fpos = fpos + 1
      Next
      FetchStatus = dr4.EOF
      If FetchStatus = False Then dr4.MoveNext
   Case 5
      For z = 1 To cfmax
         Call GetVariablePositionAndType(ColsFetch(z), pos%, ty%)
         If ty% = 1 Then
            If IsNull(dr5.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = dr5.Fields(fpos).Value
            End If
         Else
            If IsNull(dr5.Fields(fpos).Value) Then
               Vars(pos%, 3) = "<NULL>"
            Else
               Vars(pos%, 3) = Str(dr5.Fields(fpos).Value)
            End If
         End If
         fpos = fpos + 1
      Next
      FetchStatus = dr5.EOF
      If FetchStatus = False Then dr5.MoveNext
End Select

Call UpdateVariables
On Error GoTo 0
Exit Sub

errorfound:
   If FetchStatus Or InStr(LCase(Err.Description), "EOF ") Then
      MsgBox "Cursor did not pick up any records."
   End If
   Resume Next

End Sub

Public Sub GetVariablePositionAndType(var$, pos%, ty%)

For z = 1 To vmax
  If InStr(var$, Vars(z, 1)) Then
      pos% = z
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            ty% = 0
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            ty% = 1
      End Select
      Exit For
  End If
Next

End Sub

Public Sub DoEnd()

   If WhileInProgress Then
      If FetchStatus = False Then
         CurrPos = WhileStartPos
         Exit Sub
      Else
         WhileInProgress = False
      End If
   End If
   
   If IFInProgress Then
      If IFStatus(IFInProgress) Then
         z = CurrPos + 1
         flex.Row = z
         flex.Col = 1
         If LCase(Trim(flex.Text)) = "else" Then
            Do Until LCase(Trim(flex.Text)) = "end"
               z = z + 1
               flex.Row = z
            Loop
            CurrPos = z
         End If
         CurrPos = z
      End If
      IFInProgress = IFInProgress - 1
   End If

End Sub

Public Sub DoWhile()

WhileStartPos = CurrPos - 1
WhileInProgress = True
If InStr(LCase(flex.Text), "@@fetch_status") > 0 And InStr(flex.Text, "=") > 0 And InStr(flex.Text, "0") > 0 Then
   If FetchStatus Then
      flex.Col = 1
      For z = CurrPos To flex.Rows - 1
         flex.Row = z
         cs$ = Trim(LCase(flex.Text))
         If cs$ = "end" Then
            CurrPos = z
            Exit For
         End If
      Next
      WhileInProgress = False
   End If
End If

If FetchStatus Then
   flex.Col = 1
   For z = CurrPos To flex.Rows - 1
      flex.Row = z
      cs$ = Trim(LCase(flex.Text))
      If cs$ = "end" Then
         CurrPos = z
         Exit For
      End If
   Next
   WhileInProgress = False
End If

'maybe another format later?

End Sub

Public Sub DoExec()

On Error GoTo errorfound

cs$ = Trim(flex.Text)
p = InStr(LCase(cs$), "exec")
p = InStr(p + 1, cs$, " ")
cs$ = Trim(Mid(cs$, p + 1))
If InStr(cs$, "=") Then GoTo EqualCall
'replace vars
For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
      End Select
  End If
Next

dcmd1.CommandText = cs$
dcmd1.Execute
On Error GoTo 0
Exit Sub

EqualCall:
eqpos = InStr(cs$, "=")
currvar = Trim(Left(cs$, eqpos - 1))
cs$ = Trim(Mid(cs$, eqpos + 1))
For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
      End Select
  End If
Next
dcmd1.CommandText = cs$
Set drs = dcmd1.Execute
For z = 1 To vmax
   If currvar = Vars(z, 1) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            If drs.BOF Or drs.EOF Then
               Vars(z, 3) = "<NULL>"
            Else
               Vars(z, 3) = Str(drs.Fields(0).Value)
            End If
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If drs.BOF Or drs.EOF Then
               Vars(z, 3) = "<NULL>"
            Else
               If IsNull(drs.Fields(0).Value) Then
                   Vars(z, 3) = "<NULL>"
               Else
                   Vars(z, 3) = drs.Fields(0).Value
               End If
            End If
      End Select
   End If
Next
Call UpdateVariables
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next
   
End Sub

Public Sub DoPrint()

On Error GoTo errorfound

cs$ = Trim(flex.Text)
p = InStr(LCase(cs$), "print")
p = InStr(p + 1, cs$, " ")
cs$ = Trim(Mid(cs$, p + 1))
'replace vars
For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
      End Select
  End If
Next
cs$ = Replace(cs$, "'", "")
lbPrint.Caption = cs$
lbPrint.Refresh
On Error GoTo 0
Exit Sub

errorfound:
   MsgBox "There was the error - " + Err.Description
   Resume Next
   
End Sub

Public Sub DoEval(cs$)

cs$ = Trim(flex.Text)
eqpos = 0
If InStr(cs$, "=") Then eqpos = InStr(cs$, "=")
p = InStr(cs$, "@")
If p = 0 Then Exit Sub
currvar = ""
Do Until Mid(cs$, p, 1) = " "
   currvar = currvar + Mid(cs$, p, 1)
   p = p + 1
Loop

For z = 1 To vmax
  If InStr(cs$, Vars(z, 1)) And InStr(cs$, Vars(z, 1)) > eqpos Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If Vars(z, 3) = "<NULL>" Then
                cs$ = Replace(cs$, Vars(z, 1), "''")
            Else
                cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
            End If
      End Select
  Else
     p = InStr(eqpos + 1, cs$, Vars(z, 1))
     If p Then
        Select Case Vars(z, 2)
           Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
              cs$ = Left(cs$, p - 1) + Vars(z, 3) + Mid(cs$, p + Len(currvar))
           Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
              If Vars(z, 3) = "<NULL>" Then
                 cs$ = Left(cs$, p - 1) + "''" + Mid(cs$, p + Len(currvar))
              Else
                  cs$ = Left(cs$, p - 1) + "'" + Vars(z, 3) + "'" + Mid(cs$, p + Len(currvar))
              End If
        End Select
     End If
  End If
Next

Restart:
For z = 1 To 3
   If InStr(cs$, Mid(matheq, z, 1)) Then
      typev$ = Mid(matheq, z, 1)
      Exit For
   End If
Next
'do it
pos = InStr(cs$, typev$)
If Mid(cs$, pos - 1, 1) = " " Then
   leftstart = pos - 2
Else
   leftstart = pos - 1
End If
If Mid(cs$, pos + 1, 1) = " " Then
   rightstart = pos + 2
Else
   rightstart = pos + 1
End If
lvar$ = ""
rvar$ = ""

leftmost = 0
If Mid(cs$, leftstart, 1) = "'" Then
   chkchr$ = "'"
   lvar$ = "'"
   leftstart = leftstart - 1
Else
   chkchr$ = " "
End If
For z = leftstart To 1 Step -1
   If Mid(cs$, z, 1) = chkchr$ Then
      leftmost = z
      Exit For
   Else
      lvar$ = Mid(cs$, z, 1) + lvar$
   End If
Next

rightmost = 0
If Mid(cs$, rightstart, 1) = "'" Then
   chkchr$ = "'"
   rvar$ = "'"
   rightstart = rightstart + 1
Else
   chkchr$ = " "
End If
For z = rightstart To Len(cs$)
   If Mid(cs$, z, 1) = chkchr$ Then
      rightmost = z
      Exit For
   Else
      rvar$ = rvar$ + Mid(cs$, z, 1)
   End If
Next
If rightmost = 0 Then rightmost = Len(cs$)

'set leftmost and rightmost string
datatype = "N"
If chkchr$ = "'" Then datatype = "S"
Select Case typev$
   Case "+"
      If datatype = "S" Then
         nval$ = "'" + lvar$ + rvar$ + "'"
         nval$ = "'" + Replace(nval$, "'", "") + "'"
         cs$ = Replace(cs$, Mid(cs$, leftmost, (rightmost - leftmost) + 1), nval$)
      Else
         newval = Val(lvar$) + Val(rvar$)
         If newval <> 0 Then
            nval$ = Trim(Str(newval))
            cs$ = Replace(cs$, Mid(cs$, leftmost, (rightmost - leftmost) + 1), nval$)
         End If
      End If
   Case "*"
      newval = Val(lvar$) * Val(rvar$)
      If newval <> 0 Then
         nval$ = Trim(Str(newval))
         cs$ = Replace(cs$, Mid(cs$, leftmost, (rightmost - leftmost) + 1), nval$)
      End If
   Case "-"
      newval = Val(lvar$) - Val(rvar$)
      If newval <> 0 Then
         nval$ = Trim(Str(newval))
         cs$ = Replace(cs$, Mid(cs$, leftmost, (rightmost - leftmost) + 1), nval$)
      End If
   Case "/"
      If Val(rvar$) <> 0 Then
         newval = Val(lvar$) / Val(rvar$)
         If newval <> 0 Then
            nval$ = Trim(Str(newval))
            cs$ = Replace(cs$, Mid(cs$, leftmost, (rightmost - leftmost) + 1), nval$)
         End If
      End If
End Select
'now see if anymore to do
flag% = 0
For z = 1 To 3
   If InStr(cs$, Mid(matheq, z, 1)) Then flag% = 1
Next
If flag% Then GoTo Restart

If eqpos = 0 Then Exit Sub

If Left(nval$, 1) = "'" And Right(nval$, 1) = "'" Then nval$ = Replace(nval$, "'", "")
For z = 1 To vmax
   If currvar = Vars(z, 1) Then
      Select Case Vars(z, 2)
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            If nval$ = "" Then
               Vars(z, 3) = "0"
            Else
               Vars(z, 3) = nval$
            End If
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            If nval$ = "" Then
               Vars(z, 3) = "<NULL>"
            Else
                Vars(z, 3) = nval$
            End If
      End Select
   End If
Next
Call UpdateVariables
StatementDone = True

End Sub

Public Sub CheckForFunctions(cs$)

cs$ = flex.Text
OldFlexText = flex.Text
cs$ = LCase(cs$)
FunctionPresent = False
LastFunction = ""

If InStr(cs$, " left(") Or InStr(cs$, " right(") Or InStr(cs$, " substring(") Or InStr(cs$, " replace(") Or InStr(cs$, " datediff(") Then
   FunctionPresent = True
   FunctionType = 1
End If

If InStr(cs$, " upper(") Or InStr(cs$, " lower(") Or InStr(cs$, " ltrim(") Or InStr(cs$, " rtrim(") Then
   FunctionPresent = True
   FunctionType = 2
End If
If InStr(cs$, " len(") Or InStr(cs$, " char(") Or InStr(cs$, " space(") Or InStr(cs$, " month(") Or InStr(cs$, " day(") Or InStr(cs$, " year(") Then
   FunctionPresent = True
   FunctionType = 2
End If
If InStr(cs$, " convert(") Then
   FunctionPresent = True
   FunctionType = 3
End If

If InStr(cs$, " getdate()") Then
   FunctionPresent = True
   FunctionType = 4
End If

cs$ = flex.Text
If FunctionPresent = False Then Exit Sub
flag% = 1

Again:
LastFunction = ""
eqpos = 0
If InStr(cs$, "=") Then eqpos = InStr(cs$, "=")
Select Case FunctionType
   Case 1
      exc$ = ","
   Case 2
      exc$ = ")"
   Case 3
      exc$ = " ,)"
   Case Else
      exc$ = ""
End Select
For z = 1 To vmax
    For zz = 1 To Len(exc$)
        If InStr(cs$, Vars(z, 1) + Mid(exc$, zz, 1)) And InStr(cs$, Vars(z, 1)) > eqpos Then
            Select Case Vars(z, 2)
               Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
                  cs$ = Replace(cs$, Vars(z, 1), Vars(z, 3))
               Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
                  If Vars(z, 3) = "<NULL>" Then
                      cs$ = Replace(cs$, Vars(z, 1), "''")
                  Else
                      cs$ = Replace(cs$, Vars(z, 1), "'" + Vars(z, 3) + "'")
                  End If
            End Select
        Else
           p = InStr(eqpos + 1, cs$, Vars(z, 1) + Mid(exc$, zz, 1))
           If p Then
              Select Case Vars(z, 2)
                 Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
                    cs$ = Left(cs$, p - 1) + Vars(z, 3) + Mid(cs$, p + Len(currvar))
                 Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
                    If Vars(z, 3) = "<NULL>" Then
                       cs$ = Left(cs$, p - 1) + "''" + Mid(cs$, p + Len(currvar))
                    Else
                        cs$ = Left(cs$, p - 1) + "'" + Vars(z, 3) + "'" + Mid(cs$, p + Len(currvar))
                   End If
               End Select
           End If
        End If
    Next
Next

'*** GETDATE() must be first!
If InStr(LCase(cs$), "getdate()") Then
   cs$ = Replace(cs$, "getdate()", "'" + Str(Now) + "'")
   GoTo Done
End If

'If InStr(LCase(cs$), "cast(") Then
'   If InStr(cs$, "select ") = 0 Then
'       MsgBox "Cast can only be used for declared variables."
'       GoTo Done
'   Else 'do it
'   End If
'End If

If InStr(LCase(cs$), "convert(") Then
   If InStr(cs$, "select ") Then
       MsgBox "Convert can only be used for declared variables."
   End If
   LastFunction = "convert"
   p = InStr(LCase(cs$), "convert(")
   os$ = ""
DoMore:
   Do Until Mid(cs$, p, 1) = ")"
     os$ = os$ + Mid(cs$, p, 1)
     p = p + 1
   Loop
   os$ = os$ + ")"
   p = p + 1
   If Mid(cs$, p, 1) = "," Then GoTo DoMore
   numcommas = 0
   p = InStr(os$, ",")
   Erase ConvertData
CheckComm:
   If p Then
      numcommas = numcommas + 1
      ConvertData(numcommas) = Str(p)
      p = InStr(p + 1, os$, ",")
      GoTo CheckComm
   End If
   cstyle = ""
   ctype = ""
   clen = ""
   cvalue = ""
   p = InStr(LCase(os$), "convert(") + 8
   Do Until p = Val(ConvertData(1))
      ctype = ctype + Mid(os$, p, 1)
      p = p + 1
   Loop
   If Val(ConvertData(2)) Then
     p = Val(ConvertData(2)) + 1
     Do Until Mid(os$, p, 1) = ")"
         cstyle = cstyle + Mid(os$, p, 1)
         p = p + 1
     Loop
     cstyle = Trim(cstyle)
   Else
      cstyle = "100"
   End If
   p = Val(ConvertData(1)) + 1
   If Val(ConvertData(2)) Then
      sp$ = ","
   Else
      sp$ = ")"
   End If
   Do Until Mid(os$, p, 1) = sp$
      cvalue = cvalue + Mid(os$, p, 1)
      p = p + 1
   Loop
   cvalue = Trim(cvalue)
   cvalue = Replace(cvalue, "'", "")
   'type is gotten and style ready now...
   If InStr(ctype, "(") Then
      p = InStr(ctype, "(") + 1
      pos = p - 1
      Do Until Mid(ctype, p, 1) = ")"
         clen = clen + Mid(ctype, p, 1)
         p = p + 1
      Loop
      ctype = Left(ctype, pos - 1)
   End If
   Select Case ctype
         Case "bigint", "bit", "decimal", "float", "int", "money", "numeric", "real", "smallint", "smallmoney", "tinint"
            dtype = "Numeric"
         Case "char", "datetime", "nchar", "nvarchar", "smalldatetime", "text", "uniqueidentifier", "varbinary", "varchar"
            dtype = "String"
   End Select
   Select Case cstyle
     Case "0"
        sv$ = Format$(cvalue, "mm dd yy hh:mm")
     Case "1"
        sv$ = Format$(cvalue, "mm/dd/yy")
     Case "2"
        sv$ = Format$(cvalue, "yy.mm.dd")
     Case "3"
        sv$ = Format$(cvalue, "dd/mm/yy")
     Case "4"
        sv$ = Format$(cvalue, "dd.mm.yy")
     Case "5"
        sv$ = Format$(cvalue, "dd-mm-yy")
     Case "6"
        sv$ = Format$(cvalue, "dd mmm yy")
     Case "7"
        sv$ = Format$(cvalue, "mmm dd yy")
     Case "8", "108"
        sv$ = Format$(cvalue, "hh:mm:ss")
     Case "10"
        sv$ = Format$(cvalue, "mm-dd-yy")
     Case "11"
        sv$ = Format$(cvalue, "yy/mm/dd")
     Case "12"
        sv$ = Format$(cvalue, "yymmdd")
     Case "100"
        sv$ = Format$(cvalue, "mm dd yyyy hh:mm")
     Case "101"
        sv$ = Format$(cvalue, "mm/dd/yyyy")
     Case "102"
        sv$ = Format$(cvalue, "yyyy.mm.dd")
     Case "103"
        sv$ = Format$(cvalue, "dd/mm/yyyy")
     Case "104"
        sv$ = Format$(cvalue, "dd.mm.yyyy")
     Case "105"
        sv$ = Format$(cvalue, "dd-mm-yyyy")
     Case "106"
        sv$ = Format$(cvalue, "dd mmm yyyy")
     Case "107"
        sv$ = Format$(cvalue, "mmm dd yyyy")
     Case "110"
        sv$ = Format$(cvalue, "mm-dd-yyyy")
     Case "111"
        sv$ = Format$(cvalue, "yyyy/mm/dd")
     Case "112"
        sv$ = Format$(cvalue, "yyyymmdd")
   End Select
   Select Case dtype
      Case "String"
         If ctype = "char" Then sv$ = Left(sv$ + Space$(Val(clen)), Val(clen))
         cs$ = Replace(cs$, os$, "'" + sv$ + "'")
      Case "Numeric"
         cs$ = Replace(cs$, os$, sv$)
   End Select
   GoTo Done
End If

'*** CHAR
If InStr(LCase(cs$), "char(") Then
   pos = InStr(LCase(cs$), "char(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "(") + 1
   Do Until Mid(os$, p, 1) = ")"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = Chr$(Val(sv$))
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** DATEDIFF
If InStr(LCase(cs$), "datediff(") Then
   pos = InStr(LCase(cs$), "datediff(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   p = InStr(LCase(cs$), "datediff") + 9
   sv$ = ""
   ss1$ = ""
   ss2$ = ""
   Do Until Mid(cs$, p, 1) = ","
      sv$ = sv$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   sv$ = Trim(sv$)
   p = p + 1
ddloop1:
   If Mid(cs$, p, 1) <> "'" Then
      p = p + 1
      GoTo ddloop1
   End If
   p = p + 1
   Do Until Mid(cs$, p, 1) = "'"
      ss1$ = ss1$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = p + 1
ddloop2:
   If Mid(cs$, p, 1) <> "'" Then
      p = p + 1
      GoTo ddloop2
   End If
   p = p + 1
   Do Until Mid(cs$, p, 1) = "'"
      ss2$ = ss2$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   Dim d1 As Date
   Dim d2 As Date
   d1 = ss1$
   d2 = ss2$
   Select Case sv$
       Case "year", "yy", "yyyy"
           dff = DateDiff("yyyy", ss1$, ss2$)
       Case "quarter", "qq", "q"
           dff = DateDiff("q", ss1$, ss2$)
       Case "month", "mm", "m"
           dff = DateDiff("m", ss1$, ss2$)
       Case "dayofyear", "dy", "y"
           dff = DateDiff("y", ss1$, ss2$)
       Case "day", "dd", "d"
           dff = DateDiff("d", ss1$, ss2$)
       Case "week", "wk", "w"
           dff = DateDiff("ww", ss1$, ss2$)
       Case "hour", "hh"
           dff = DateDiff("h", ss1$, ss2$)
       Case "minute", "mi", "n"
           dff = DateDiff("n", ss1$, ss2$)
       Case "second", "ss", "s"
           dff = DateDiff("s", ss1$, ss2$)
   End Select
   sv$ = Trim(Str(dff))
   cs$ = Replace(cs$, os$, sv$)
   GoTo Done
End If

'*** DAY
If InStr(LCase(cs$), "day(") Then
   pos = InStr(LCase(cs$), "day(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = Day(sv$)
   cs$ = Replace(cs$, os$, sv$)
   GoTo Done
End If

'*** LEFT
If InStr(LCase(cs$), "left(") Then
   pos = InStr(LCase(cs$), "left(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   p = InStr(InStr(LCase(cs$), "left"), cs$, "'") + 1
   sv$ = ""
   ss$ = ""
   Do Until Mid(cs$, p, 1) = "'"
      sv$ = sv$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = InStr(p, cs$, ",")
   p = p + 1
   Do Until Mid(cs$, p, 1) = ")"
      ss$ = ss$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   sv$ = Left(sv$, Val(ss$))
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** LEN
If InStr(LCase(cs$), "len(") Then
   pos = InStr(LCase(cs$), "len(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   slv = Len(sv$)
   cs$ = Replace(cs$, os$, Str(slv))
   GoTo Done
End If

'*** LOWER
If InStr(LCase(cs$), "lower(") Then
   pos = InStr(LCase(cs$), "lower(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = LCase(sv$)
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** LTRIM
If InStr(LCase(cs$), "ltrim(") Then
   pos = InStr(LCase(cs$), "ltrim(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = LTrim(sv$)
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'***MONTH
If InStr(LCase(cs$), "month(") Then
   pos = InStr(LCase(cs$), "month(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = Month(sv$)
   cs$ = Replace(cs$, os$, sv$)
   GoTo Done
End If

'***REPLACE
If InStr(LCase(cs$), "replace(") Then
   pos = InStr(LCase(cs$), "replace(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   p = InStr(InStr(LCase(cs$), "replace"), cs$, "'") + 1
   sv$ = ""
   ss1$ = ""
   ss2$ = ""
   Do Until Mid(cs$, p, 1) = "'"
      sv$ = sv$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = InStr(p, cs$, ",")
   p = p + 1
reploop1:
   If Mid(cs$, p, 1) <> "'" Then
      p = p + 1
      GoTo reploop1
   End If
   p = p + 1
   Do Until Mid(cs$, p, 1) = "'"
      ss1$ = ss1$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = p + 1
reploop2:
   If Mid(cs$, p, 1) <> "'" Then
      p = p + 1
      GoTo reploop2
   End If
   p = p + 1
   Do Until Mid(cs$, p, 1) = "'"
      ss2$ = ss2$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   sv$ = Replace(sv$, ss1$, ss2$)
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** RIGHT
If InStr(LCase(cs$), "right(") Then
   pos = InStr(LCase(cs$), "right(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   p = InStr(InStr(LCase(cs$), "right"), cs$, "'") + 1
   sv$ = ""
   ss$ = ""
   Do Until Mid(cs$, p, 1) = "'"
      sv$ = sv$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = InStr(p, cs$, ",")
   p = p + 1
   Do Until Mid(cs$, p, 1) = ")"
      ss$ = ss$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   sv$ = Right(sv$, Val(ss$))
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** RTRIM
If InStr(LCase(cs$), "rtrim(") Then
   pos = InStr(LCase(cs$), "rtrim(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = RTrim(sv$)
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** SPACE
If InStr(LCase(cs$), "space(") Then
   pos = InStr(LCase(cs$), "space(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "(") + 1
   Do Until Mid(os$, p, 1) = ")"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = Space$(Val(sv$))
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'***SUBSTRING
If InStr(LCase(cs$), "substring(") Then
   pos = InStr(LCase(cs$), "substring(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   p = InStr(InStr(LCase(cs$), "substring"), cs$, "'") + 1
   sv$ = ""
   ss1$ = ""
   ss2$ = ""
   Do Until Mid(cs$, p, 1) = "'"
      sv$ = sv$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = InStr(p, cs$, ",")
   p = p + 1
   Do Until Mid(cs$, p, 1) = ","
      ss1$ = ss1$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   p = p + 1
   Do Until Mid(cs$, p, 1) = ")"
      ss2$ = ss2$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   sv$ = Mid(sv$, Val(ss1$), Val(ss2$))
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** UPPER
If InStr(LCase(cs$), "upper(") Then
   pos = InStr(LCase(cs$), "upper(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = UCase(sv$)
   cs$ = Replace(cs$, os$, "'" + sv$ + "'")
   GoTo Done
End If

'*** YEAR
If InStr(LCase(cs$), "year(") Then
   pos = InStr(LCase(cs$), "year(")
   os$ = ""
   p = pos
   Do Until Mid(cs$, p, 1) = ")"
      os$ = os$ + Mid(cs$, p, 1)
      p = p + 1
   Loop
   os$ = os$ + ")"
   sv$ = ""
   p = InStr(os$, "'") + 1
   Do Until Mid(os$, p, 1) = "'"
      sv$ = sv$ + Mid(os$, p, 1)
      p = p + 1
   Loop
   sv$ = Year(sv$)
   cs$ = Replace(cs$, os$, sv$)
   GoTo Done
End If

Done:
flex.Text = cs$
If InStr(cs$, " left(") Or InStr(cs$, " right(") Or InStr(cs$, " substring(") Or InStr(cs$, " replace(") Or InStr(cs$, " datediff(") Then
   GoTo Again
End If
If InStr(cs$, " upper(") Or InStr(cs$, " lower(") Or InStr(cs$, " ltrim(") Or InStr(cs$, " rtrim(") Or InStr(cs$, " len(") Or InStr(cs$, " char(") Or InStr(cs$, " space(") Then
   GoTo Again
End If
If InStr(cs$, " month(") Or InStr(cs$, " day(") Or InStr(cs$, " year(") Then
   GoTo Again
End If
If InStr(cs$, " getdate()") Or InStr(cs$, " convert(") Then
   GoTo Again
End If

End Sub
