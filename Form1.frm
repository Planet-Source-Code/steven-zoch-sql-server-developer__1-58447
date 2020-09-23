VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "SQL Admin"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkSys 
      Caption         =   "sys"
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkDT 
      Caption         =   "dt's"
      Height          =   255
      Left            =   8400
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   2520
   End
   Begin VB.ListBox List5 
      Height          =   2790
      Left            =   12000
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox List4 
      Height          =   2790
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   360
      Width           =   3495
   End
   Begin VB.ListBox List3 
      Height          =   2790
      Left            =   4560
      TabIndex        =   12
      Top             =   360
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   14775
      Begin VB.CommandButton Command5 
         Caption         =   ">"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid flex 
         Height          =   4935
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   8705
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   14775
      Begin VB.CommandButton Command6 
         Caption         =   "Clear All"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   800
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   120
         Visible         =   0   'False
         Width           =   12975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Execute (F5)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   13335
      End
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   10080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Triggers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   17
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Stored Procedures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   15
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lbStat 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   10080
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Columns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu fle 
      Caption         =   "&File"
      Begin VB.Menu opt 
         Caption         =   "&Options..."
      End
      Begin VB.Menu flemod 
         Caption         =   "Modify Servers"
         Begin VB.Menu flecl 
            Caption         =   "Create New Server Link..."
         End
         Begin VB.Menu flemp 
            Caption         =   "Modify User/Passwords..."
         End
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu fleext 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu db 
      Caption         =   "Database"
      Begin VB.Menu dbcreate 
         Caption         =   "Create Database..."
      End
      Begin VB.Menu dbdel 
         Caption         =   "Delete Database"
      End
      Begin VB.Menu dbdep 
         Caption         =   "Dependencies"
      End
      Begin VB.Menu dbdiv1 
         Caption         =   "-"
      End
      Begin VB.Menu dbvwinx 
         Caption         =   "View Indexes"
         Begin VB.Menu vwdb 
            Caption         =   "All Databases"
         End
         Begin VB.Menu vwdbc 
            Caption         =   "Current Database"
         End
      End
   End
   Begin VB.Menu scr 
      Caption         =   "Scripting"
      Begin VB.Menu scrdb 
         Caption         =   "Script Database..."
      End
      Begin VB.Menu scrtb 
         Caption         =   "Script Table..."
      End
      Begin VB.Menu scrproc 
         Caption         =   "Process Script..."
      End
   End
   Begin VB.Menu ser 
      Caption         =   "Search"
   End
   Begin VB.Menu sproc 
      Caption         =   "StoredProcedures"
      Begin VB.Menu spcret 
         Caption         =   "Create..."
      End
      Begin VB.Menu spdebug 
         Caption         =   "Debug..."
      End
   End
   Begin VB.Menu tb 
      Caption         =   "Tables"
      Begin VB.Menu alttb 
         Caption         =   "Alter Table..."
      End
      Begin VB.Menu cpytb 
         Caption         =   "Copy Table As..."
      End
      Begin VB.Menu tbcreate 
         Caption         =   "Create New Table..."
      End
      Begin VB.Menu tbdep 
         Caption         =   "Dependencies"
      End
      Begin VB.Menu indx 
         Caption         =   "Indexes"
         Begin VB.Menu indxcr 
            Caption         =   "Create Index..."
         End
         Begin VB.Menu indxdr 
            Caption         =   "Drop Index..."
         End
         Begin VB.Menu indvw 
            Caption         =   "View"
         End
      End
   End
   Begin VB.Menu utl 
      Caption         =   "Utilities"
      Begin VB.Menu utlalt 
         Caption         =   "Alter Screen"
         Begin VB.Menu altrea 
            Caption         =   "Really Squeeze"
         End
         Begin VB.Menu altsq 
            Caption         =   "Squeeze"
         End
         Begin VB.Menu altunsq 
            Caption         =   "Unsqueeze"
         End
      End
      Begin VB.Menu tbclr 
         Caption         =   "Clear Table"
      End
      Begin VB.Menu utlmv 
         Caption         =   "Copy"
         Begin VB.Menu utlsp 
            Caption         =   "Stored Procedures..."
         End
         Begin VB.Menu cpytrig 
            Caption         =   "Triggers..."
         End
      End
      Begin VB.Menu utldel 
         Caption         =   "Delete..."
      End
      Begin VB.Menu dep 
         Caption         =   "Dependencies"
         Begin VB.Menu depstr 
            Caption         =   "Stored Procedure"
         End
         Begin VB.Menu deptrig 
            Caption         =   "Triggers"
         End
      End
      Begin VB.Menu utlpnt 
         Caption         =   "Print Screen"
      End
      Begin VB.Menu utlren 
         Caption         =   "Rename..."
      End
   End
   Begin VB.Menu hp 
      Caption         =   "Help"
      Begin VB.Menu hpabt 
         Caption         =   "About Sql Developer"
      End
   End
   Begin VB.Menu popmenu 
      Caption         =   "popmnu"
      Visible         =   0   'False
      Begin VB.Menu popdel 
         Caption         =   "Delete Col (Display only)"
      End
      Begin VB.Menu popedt 
         Caption         =   "Edit"
      End
      Begin VB.Menu popins 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu popdiv 
         Caption         =   "-"
      End
      Begin VB.Menu popext 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StMax As Integer
Public oldtext As String
Public keystr As String

Private Sub altrea_Click()

List1.Height = 1230
List2.Height = 1230
List3.Height = 1230
List4.Height = 1230
List5.Height = 1230

Frame2.Height = 8535
Frame2.Top = 1560

flex.Height = 8055
flex.Top = 360

End Sub

Private Sub altsq_Click()

List1.Height = 1230
List2.Height = 1230
List3.Height = 1230
List4.Height = 1230
List5.Height = 1230

Frame1.Top = 1560

Frame2.Height = 6975
Frame2.Top = 3120

flex.Height = 6495
flex.Top = 360

End Sub

Private Sub alttb_Click()

If TableName = "" Then
   MsgBox "You must first select a Table."
   Exit Sub
End If

frmAlter.Show 1
List2_Click

End Sub

Private Sub altunsq_Click()

List1.Height = 2790
List2.Height = 2790
List3.Height = 2790
List4.Height = 2790
List5.Height = 2790

Frame1.Top = 3120

Frame2.Height = 5535
Frame2.Top = 4560

flex.Height = 4935
flex.Top = 360

End Sub

Private Sub chkDT_Click()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = "sp_help"
Set rs = cmd1.Execute
rs.Filter = "Object_type='stored procedure'"
List4.Clear
keystr = ""
Do While Not rs.EOF
    If chkDT.Value = 0 Then
       If Left(rs!Name, 3) <> "dt_" Then List4.AddItem rs!Name
    Else
       List4.AddItem rs!Name
    End If
    rs.MoveNext
Loop
conn.Close

End Sub
Private Sub chkSys_Click()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Set rs = conn.OpenSchema(adSchemaTables)
List2.Clear
keystr = ""
Do While Not rs.EOF
    If chkSys.Value = 0 Then
       If UCase(Left(rs!table_name, 3)) <> "SYS" Then List2.AddItem rs!table_name
    Else
       List2.AddItem rs!table_name
    End If
    rs.MoveNext
Loop

End Sub

Private Sub Combo1_Click()

Text1.Text = Combo1.Text
Combo1.Visible = False

End Sub

Private Sub Command1_Click()

Dim CurrState(500) As String

If Text1.Text = "" Then Exit Sub
Text1.Text = Text1.Text + vbCrLf
CurrentDBName = DBName
numstate% = 0
numselect% = 0
'determine number of statements to perform
ss$ = ""
t$ = Text1.Text
If InStr(LCase(t$), "create procedure") Or InStr(LCase(t$), "create trigger") Or InStr(LCase(t$), "alter procedure") Or InStr(LCase(t$), "alter trigger") Then
   numstate% = 1
   CurrState(numstate%) = t$
Else
   For z = 1 To Len(t$)
      Select Case Mid(t$, z, 1)
         Case Chr(13)
            numstate% = numstate% + 1
            CurrState(numstate%) = ss$
         Case Chr(10)
            ss$ = ""
         Case Else
            ss$ = ss$ + Mid(t$, z, 1)
      End Select
   Next
End If

flag% = 0
For z = 1 To StMax
   If Text1.Text = Statements(z) Then
      flag% = 1
      Exit For
   End If
Next
If flag% = 0 Then
   StMax = StMax + 1
   Statements(StMax) = Text1.Text
   Combo1.Clear
   For z = 1 To StMax
      Combo1.AddItem Statements(z)
   Next
   Combo1.Text = Combo1.List(0)
End If

StopNow = False
Command2.Enabled = True
frmMain.MousePointer = 11
flex.Redraw = False
flex.Visible = False
Dim sizes(500) As Integer
Dim smax As Integer
On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

CountSqlStr = CountSql
If InStr(LCase(Text1.Text), "where ") Then
  CountSqlStr = CountSqlStr + " " + Mid(Text1.Text, InStr(LCase(Text1.Text), "where "))
End If
If InStr(LCase(CountSqlStr), "order by") Then
   CountSqlStr = Left(CountSqlStr, InStr(LCase(CountSqlStr), "order by") - 1)
End If
cmd1.CommandText = CountSqlStr
If CountSqlStr <> "" And CountSql <> "" Then
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      frmMain.MousePointer = 1
      MsgBox "There were errors or no data returned."
      GoTo Done
   End If
   reccount = rs!cmax
End If

For stcnt = 1 To numstate%
  cstate$ = CurrState(stcnt)
  If cstate$ = "" Then GoTo Done
  If Left(LCase(cstate$), 7) = "select " And numselect% Then
     CurrentStatement = cstate$
     'find next available window
     If Select01InUse = False Then
        frmSelect.Show
        GoTo Done
     End If
     If Select02InUse = False Then
        frmSelect2.Show
        GoTo Done
     End If
     If Select03InUse = False Then
        frmSelect3.Show
        GoTo Done
     End If
     If Select04InUse = False Then
        frmSelect4.Show
        GoTo Done
     End If
     If Select05InUse = False Then
        frmSelect5.Show
        GoTo Done
     End If
     If Select06InUse = False Then
        frmSelect6.Show
        GoTo Done
     End If
     If Select07InUse = False Then
        frmSelect7.Show
        GoTo Done
     End If
     If Select08InUse = False Then
        frmSelect8.Show
        GoTo Done
     End If
     If Select09InUse = False Then
        frmSelect9.Show
        GoTo Done
     End If
     If Select10InUse = False Then
        frmSelect10.Show
        GoTo Done
     End If
     MsgBox "No more than ten windows can be opened at one time."
     GoTo Done
  End If

TryAgain:
cmd1.CommandText = cstate$
If Left(LCase(cstate$), 5) = "drop " Then
  dcheck& = CheckDependencies(cstate$, msg$)
  If dcheck& Then
     sysbut% = MsgBox("The following dependencies are connected to this object:" + vbCrLf + msg$ + vbCrLf + "Are you sure you want to drop it?", 4, "Drop confirm")
     If sysbut% <> vbYes Then GoTo Done
  End If
End If

If Left(LCase(cstate$), 7) = "select " Then
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      frmMain.MousePointer = 1
      MsgBox "There were errors or no data returned."
      GoTo Done
   End If
   RecordSetReturned = True
Else
   stat& = DetermineIfNotSQLStatement(cstate$)
   If stat& Then
      Set rs = cmd1.Execute
      RecordSetReturned = True
   Else
      cmd1.Execute
      RecordSetReturned = True
   End If
End If

If Left(LCase(cstate$), 5) = "drop " Then
   If InStr(LCase(cstate$), "table ") Or InStr(LCase(cstate$), "procedure ") Or InStr(LCase(cstate$), "trigger ") Then List1_Click
   If InStr(LCase(cstate$), "database ") Then Form_Load
   LogAction = "Drop Performed"
   LogComment = cstate$ + " in database " + DBName
   Call UpdateLog
End If

If Left(LCase(cstate$), 4) = "use " Then
   CurrentDBName = Mid(cstate$, 5)
   RecordSetReturned = False
End If

If RecordSetReturned = False Then GoTo Done
If rs.Fields.Count = 0 Then
   MsgBox "There was no data returned."
   GoTo Done
End If

PBar.Value = 0
PBar.Visible = True

'do headers
flex.Cols = rs.Fields.Count + 1
frmView!flex2.Cols = flex.Cols
flex.Row = 0
frmView!flex2.Row = 0
smax = 0
cr& = 0
For z = 0 To rs.Fields.Count - 1
   flex.Col = z + 1
   frmView!flex2.Col = z + 1
   flex.CellFontName = TFontName
   flex.CellFontSize = TFontSize
   frmView!flex2.CellFontName = TFontName
   frmView!flex2.CellFontSize = TFontSize
   flex.Text = LCase(rs.Fields(z).Name)
   frmView!flex2.Text = LCase(rs.Fields(z).Name)
   smax = smax + 1
   sizes(smax) = Len(flex.Text)
Next
'do data
Do While Not rs.EOF
   DoEvents
DoStopCheck:
   If StopNow Then Exit Do
   cr& = cr& + 1
   lbStat.Caption = "Reading" + Str(cr&)
   lbStat.Refresh
   If reccount Then
      If Int((cr& / reccount) * 100) <= 100 Then
         PBar.Value = Int((cr& / reccount) * 100)
         PBar.Refresh
      End If
   End If
   flex.Rows = cr& + 1
   flex.Row = cr&
   frmView!flex2.Rows = cr& + 1
   frmView!flex2.Row = cr&
   For z = 0 To rs.Fields.Count - 1
      flex.Col = z + 1
      frmView!flex2.Col = flex.Col
      Select Case rs.Fields(z).Type
         Case 3, 5, 131, 2, 6, 17, 4, 20 ', 130
            If IsNull(rs.Fields(z).Value) Then
               flex.Text = " 0"
               frmView!flex2.Text = " 0"
            Else
               flex.Text = " " + Trim(Str(rs.Fields(z).Value))
               frmView!flex2.Text = " " + Trim(Str(rs.Fields(z).Value))
            End If
         Case 204, 128
            acnv = StrConv(rs.Fields(z).Value, vbUnicode)
            If IsNull(acnv) Then
               hxstr = ""
            Else
               hxstr = " 0x"
               For y = 1 To Len(acnv)
                  hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
               Next
            End If
            flex.Text = " " + Trim(hxstr)
            frmView!flex2.Text = " " + Trim(hxstr)
         Case 129, 200, 135, 202, 203, 11, 72, 201, 205, 130
            If IsNull(rs.Fields(z).Value) Then
               flex.Text = ""
               frmView!flex2.Text = ""
               If NullReq = "Y" Then
                  flex.Text = "Null"
                  frmView!flex2.Text = "Null"
               End If
            Else
               flex.Text = rs.Fields(z).Value
               frmView!flex2.Text = rs.Fields(z).Value
               If IsNumeric(Left(rs.Fields(z).Value, 1)) Then
                  flex.Text = " " + flex.Text
                  frmView!flex2.Text = " " + flex.Text
               End If
            End If
         Case Else
            MsgBox "Unable to resolve type " + Str(rs.Fields(z).Type) + " on column " + rs.Fields(z).Name
      End Select
      If Len(flex.Text) > sizes(z + 1) Then sizes(z + 1) = Len(flex.Text)
   Next
   rs.MoveNext
Loop

PBar.Value = 0
PBar.Visible = False
Command2.Enabled = False

flex.ColWidth(0) = 0
frmView!flex2.ColWidth(0) = 0
For z = 1 To smax
   flex.ColWidth(z) = (sizes(z) + 1) * 100
   frmView!flex2.ColWidth(z) = (sizes(z) + 1) * 100
Next
flex.Redraw = True
flex.Visible = True
numselect% = 1

Done:
Next
frmMain.MousePointer = 1
If cr& Then
   If cr& = 1 Then
      lbStat.Caption = "One Row returned."
      lbStat.Refresh
   Else
      lbStat.Caption = Format$(cr&, "###,###,##0") + " Rows returned."
      lbStat.Refresh
   End If
Else
   lbStat.Caption = "No Rows were returned."
   lbStat.Refresh
End If
conn.Close

On Error GoTo 0
Exit Sub

errorfound:
   If Err.Number = -2147217913 Then
      Resume TryAgain
   End If
   If Err.Number <> 6 And Err.Number <> 3704 And Err.Number <> 92 And Err.Number <> 30009 And Err.Number <> 30006 And Err.Number <> 424 Then MsgBox "There was the following error - " + Err.Description
   If Err.Number = 30006 Then
      StopNow = True
      MsgBox "Cannot load all records returned.  Exiting with available memory."
      Resume DoStopCheck
   End If
   Resume Next

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Combo1.Visible = False

End Sub

Private Sub Command2_Click()

StopNow = True

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Combo1.Visible = False

End Sub

Private Sub Command3_Click()

Text1.Text = ""

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Combo1.Visible = False

End Sub

Private Sub Command4_Click()

For z = 1 To flex.Cols - 1
   If flex.ColWidth(z) - 50 > 0 Then
      flex.ColWidth(z) = flex.ColWidth(z) - 50
   End If
Next

End Sub

Private Sub Command5_Click()

For z = 1 To flex.Cols - 1
   flex.ColWidth(z) = flex.ColWidth(z) + 50
Next

End Sub

Private Sub Command6_Click()

On Error Resume Next
Text1.Text = ""
Unload frmSelect
Unload frmSelect2
Unload frmSelect3
Unload frmSelect4
Unload frmSelect5
Unload frmSelect6
Unload frmSelect7
Unload frmSelect8
Unload frmSelect9
Unload frmSelect10
Select01InUse = False
Select02InUse = False
Select03InUse = False
Select04InUse = False
Select05InUse = False
Select06InUse = False
Select07InUse = False
Select08InUse = False
Select09InUse = False
Select10InUse = False
flex.Visible = False
On Error GoTo 0

End Sub

Private Sub cpytb_Click()

If DBName = "" Then
   MsgBox "You must first select a database."
   Exit Sub
End If

frmCopyTableAs.Show 1

End Sub

Private Sub cpytrig_Click()

If List5.ListCount = 0 Then
   MsgBox "No Stored Procedures at this time."
   Exit Sub
End If

frmMoveTriggers.Show 1

End Sub

Private Sub dbcreate_Click()

frmNewDB.Show 1
Form_Load

End Sub

Private Sub dbdel_Click()

If DBName = "" Then
   MsgBox "A Database has not been selected."
   Exit Sub
End If

sysbut% = MsgBox("Are you sure you want to delete the database " + DBName + "?", 4, "Drop Confirm")
If sysbut% <> vbYes Then Exit Sub

On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "drop database " + DBName
cmd1.CommandText = Sql
cmd1.Execute
conn.Close
LogAction = "Delete Database"
LogComment = DBName
Call UpdateLog
On Error GoTo 0
Form_Load
Exit Sub

errorfound:
   MsgBox "Error - " + Err.Description
   Resume Next

End Sub

Private Sub dbdep_Click()

frmMain.MousePointer = 11
curDB = DBName
Text1.Text = ""
d$ = ""
numdep& = 0
On Error GoTo errorfound
sp = 0
ep = List1.ListCount - 1
For z = sp To ep
   If List1.Selected(z) Then
      sp = z
      ep = z
      Exit For
   End If
Next

For z = sp To ep
   List1.Selected(z) = True
   List1_Click
   Set conn = New ADODB.Connection
   Set cmd1 = New ADODB.Command
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + List1.List(z) + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   d$ = d$ + "Database " + List1.List(z) + vbCrLf
   For zz = 0 To List2.ListCount - 1
      d$ = d$ + Chr(9) + "Table " + List2.List(zz) + vbCrLf
      Sql = "sp_depends " + List2.List(zz)
      cmd1.CommandText = Sql
      Set rs = cmd1.Execute
      If rs.BOF Or rs.EOF Then
         d$ = d$ + Chr(9) + Chr(9) + "No Dependencies." + vbCrLf
      Else
         Do While Not rs.EOF
            d$ = d$ + Chr(9) + Chr(9) + rs!Name + " - " + rs!Type + vbCrLf
            numdep& = numdep& + 1
            rs.MoveNext
         Loop
      End If
   Next
   conn.Close
   List1.Selected(z) = False
Next

For z = 0 To List1.ListCount - 1
   If List1.List(z) = curDB Then
      List1.Selected(z) = True
      List1_Click
   End If
Next
On Error GoTo 0
Text1.Text = d$
frmMain.MousePointer = 1
frmText.Show 1
Text1.Text = ""

If numdep& Then
   If numdep& = 1 Then
      lbStat.Caption = "One Dependency returned."
      lbStat.Refresh
   Else
      lbStat.Caption = Format$(numdep&, "###,###,##0") + " Dependencies returned."
      lbStat.Refresh
   End If
Else
   lbStat.Caption = "No Dependencies were returned."
   lbStat.Refresh
End If
Exit Sub

errorfound:
   Debug.Print Err.Number, Sql, Err.Description
   Resume Next
   
End Sub

Private Sub depstr_Click()

If DBName = "" Then
   MsgBox "A Database must first be selected."
   Exit Sub
End If

frmMain.MousePointer = 11
Text1.Text = ""
d$ = ""
numdep& = 0
On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
sp = 0
ep = List4.ListCount - 1
For z = sp To ep
   If List4.Selected(z) Then
      sp = z
      ep = z
      Exit For
   End If
Next

For z = sp To ep
   d$ = d$ + "Stored Procedure " + List4.List(z) + vbCrLf
   Sql = "sp_depends " + List4.List(z)
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      d$ = d$ + Chr(9) + "No Dependencies." + vbCrLf
   Else
      Do While Not rs.EOF
         d$ = d$ + Chr(9) + rs!Name + " - " + rs!Type + " column-" + rs!Column + vbCrLf
         numdep& = numdep& + 1
         rs.MoveNext
      Loop
   End If
Next
conn.Close

On Error GoTo 0
Text1.Text = d$
frmMain.MousePointer = 1
frmText.Show 1
Text1.Text = ""
If numdep& Then
   If numdep& = 1 Then
      lbStat.Caption = "One Dependency returned."
      lbStat.Refresh
   Else
      lbStat.Caption = Format$(numdep&, "###,###,##0") + " Dependencies returned."
      lbStat.Refresh
   End If
Else
   lbStat.Caption = "No Dependencies were returned."
   lbStat.Refresh
End If
Exit Sub

errorfound:
   Resume Next

End Sub

Private Sub deptrig_Click()

If DBName = "" Then
   MsgBox "A Database must first be selected."
   Exit Sub
End If

frmMain.MousePointer = 11
Text1.Text = ""
d$ = ""
numdep& = 0
On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
sp = 0
ep = List5.ListCount - 1
For z = sp To ep
   If List5.Selected(z) Then
      sp = z
      ep = z
      Exit For
   End If
Next

For z = sp To ep
   d$ = d$ + "Trigger " + List5.List(z) + vbCrLf
   Sql = "sp_depends " + List5.List(z)
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      d$ = d$ + Chr(9) + "No Dependencies." + vbCrLf
   Else
      Do While Not rs.EOF
         d$ = d$ + Chr(9) + rs!Name + " - " + rs!Type + " column-" + rs!Column + vbCrLf
         numdep& = numdep& + 1
         rs.MoveNext
      Loop
   End If
Next
conn.Close

On Error GoTo 0
Text1.Text = d$
frmMain.MousePointer = 1
frmText.Show 1
Text1.Text = ""
If numdep& Then
   If numdep& = 1 Then
      lbStat.Caption = "One Dependency returned."
      lbStat.Refresh
   Else
      lbStat.Caption = Format$(numdep&, "###,###,##0") + " Dependencies returned."
      lbStat.Refresh
   End If
Else
   lbStat.Caption = "No Dependencies were returned."
   lbStat.Refresh
End If
Exit Sub

errorfound:
   Resume Next


End Sub

Private Sub flecl_Click()

frmModServers.Show 1

End Sub

Private Sub fleext_Click()

End

End Sub

Private Sub flemp_Click()

frmModServersUP.Show 1

End Sub

Private Sub flex_DblClick()

frmView.Show 1

End Sub

Private Sub flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

CurrRow = flex.MouseRow
CurrCol = flex.MouseCol
If CurrRow = 0 Then Exit Sub

If Button = 2 Then
   PopupMenu popmenu
End If

End Sub

Private Sub Form_Load()

Restart:
Open "c:\SqlDeveloperOptions.ini" For Random As #1 Len = 1
l& = LOF(1)
Close
If l& = 0 Then
   frmSetup.Show 1
   GoTo Restart
End If

LogUser = Space$(16)
NameSize = Len(LogUser)
x = GetComputerName(LogUser, NameSize)
LogUser = Replace(LogUser, "-", "")
LogUser = Left(LogUser, InStr(LogUser, Chr(0)) - 1)
lbStat.Caption = "Ready..."
lbStat.Refresh

currentDB = ""
currentTable = ""
currentSP = ""
Dflag% = 0
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      Dflag% = 1
      currentDB = List1.List(z)
      Exit For
   End If
Next
Tflag% = 0
For z = 0 To List2.ListCount - 1
   If List2.Selected(z) Then
      Tflag% = 1
      currentTable = List2.List(z)
      Exit For
   End If
Next
Sflag% = 0
For z = 0 To List4.ListCount - 1
   If List4.Selected(z) Then
      Sflag% = 1
      currentSP = List4.List(z)
      Exit For
   End If
Next

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
keystr = ""
PBar.Visible = False
Command2.Enabled = False
flex.Visible = False

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
   Line Input #1, d$
   xmax = Val(d$)
   Erase Xtras
   For z = 1 To xmax
      Line Input #1, Xtras(z)
   Next
   Line Input #1, LogServer
   Line Input #1, DBLog
   Close
End If
frmMain.Caption = "SQL Developer (Server " + Server + ")"
Text1.ForeColor = TForeColor
Text1.BackColor = TBackColor
Text1.FontName = TFontName
Text1.FontSize = TFontSize
List1.BackColor = List1TBackColor
List1.ForeColor = List1TForeColor
List2.BackColor = List2TBackColor
List2.ForeColor = List2TForeColor
List3.BackColor = List3TBackColor
List3.ForeColor = List3TForeColor
List4.BackColor = List4TBackColor
List4.ForeColor = List4TForeColor
List5.BackColor = List5TBackColor
List5.ForeColor = List5TForeColor

Call CheckForLogDeletions

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Sql = "sp_databases"
cmd1.CommandText = Sql
Set rs = cmd1.Execute

Do While Not rs.EOF
   List1.AddItem rs!Database_Name
   rs.MoveNext
Loop
conn.Close
If Dflag% Then
  For z = 0 To List1.ListCount - 1
     If List1.List(z) = currentDB Then
        List1.Selected(z) = True
        List1_Click
        Exit For
     End If
  Next
End If
If Tflag% Then
  For z = 0 To List2.ListCount - 1
     If List2.List(z) = currentTable Then
        List2.Selected(z) = True
        List2_Click
        Exit For
     End If
  Next
End If
If Sflag% Then
  For z = 0 To List4.ListCount - 1
     If List4.List(z) = currentSP Then
        List4.Selected(z) = True
        Exit For
     End If
  Next
End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next
Unload frmSelect
Unload frmSelect2
Unload frmSelect3
Unload frmSelect4
Unload frmSelect5
Unload frmSelect6
Unload frmSelect7
Unload frmSelect8
Unload frmSelect9
Unload frmSelect10
Select01InUse = False
Select02InUse = False
Select03InUse = False
Select04InUse = False
Select05InUse = False
Select06InUse = False
Select07InUse = False
Select08InUse = False
Select09InUse = False
Select10InUse = False
Close

If IFSPMax Then
   For z = 1 To IFSPMax
      Set conn = New ADODB.Connection
      Set cmd1 = New ADODB.Command
      strConnect = "Provider=SQLOLEDB;server=" + IFStoredProc(z, 1) + ";database=" + IFStoredProc(z, 2) + ";uid=" + UID + ";pwd=" + PWD + ";"
      conn.Open strConnect
      cmd1.ActiveConnection = conn
      Sql = "drop procedure sp_AdminIFCheck"
      cmd1.CommandText = Sql
      cmd1.Execute
      conn.Close
   Next
End If

End

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Combo1.Visible = False

End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If flex.Visible Then Exit Sub
If Button = 2 Then
   PopupMenu popmenu
End If

End Sub

Private Sub hpabt_Click()

frmAbout.Show 1

End Sub

Private Sub indvw_Click()

Dim sizes(500) As Integer
Dim smax As Integer

If TableName = "" Then
   MsgBox "You must first select a Table."
   Exit Sub
End If

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "sp_helpindex " + TableName
cmd1.CommandText = Sql
Set rs = cmd1.Execute
If rs.Fields.Count = 0 Then
   MsgBox "No Indexes for Table " + TableName
   Exit Sub
End If

flex.Redraw = False
flex.Visible = False
flex.Cols = rs.Fields.Count + 1
frmView!flex2.Cols = flex.Cols
flex.Row = 0
frmView!flex2.Row = 0
smax = 0
cr& = 0
For z = 0 To rs.Fields.Count - 1
   flex.Col = z + 1
   frmView!flex2.Col = z + 1
   flex.Text = LCase(rs.Fields(z).Name)
   frmView!flex2.Text = LCase(rs.Fields(z).Name)
   smax = smax + 1
   sizes(smax) = Len(flex.Text)
Next
'do data
Do While Not rs.EOF
   DoEvents
   If StopNow Then Exit Do
   cr& = cr& + 1
   If reccount Then
      If Int((cr& / reccount) * 100) <= 100 Then
         PBar.Value = Int((cr& / reccount) * 100)
         PBar.Refresh
      End If
   End If
   flex.Rows = cr& + 1
   flex.Row = cr&
   frmView!flex2.Rows = cr& + 1
   frmView!flex2.Row = cr&
   For z = 0 To rs.Fields.Count - 1
      flex.Col = z + 1
      frmView!flex2.Col = flex.Col
      Select Case rs.Fields(z).Type
         Case 3, 5, 131, 2, 6, 17, 4
            If IsNull(rs.Fields(z).Value) Then
               flex.Text = "0"
               frmView!flex2.Text = "0"
            Else
               flex.Text = Trim(Str(rs.Fields(z).Value))
               frmView!flex2.Text = Trim(Str(rs.Fields(z).Value))
            End If
         Case 204, 128
            acnv = StrConv(trs.Fields(z).Value, vbUnicode)
            If IsNull(acnv) Then
               hxstr = ""
            Else
               hxstr = "0x"
               For y = 1 To Len(acnv)
                  hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
               Next
            End If
            flex.Text = Trim(hxstr)
            frmView!flex2.Text = Trim(hxstr)
         Case 129, 200, 135, 202, 203, 11, 72, 201, 205
            If IsNull(rs.Fields(z).Value) Then
               flex.Text = ""
               frmView!flex2.Text = ""
            Else
               flex.Text = rs.Fields(z).Value
               frmView!flex2.Text = rs.Fields(z).Value
            End If
         Case Else
            MsgBox "Unable to resolve type " + Str(rs.Fields(z).Type) + " on column " + rs.Fields(z).Name
      End Select
      If Len(flex.Text) > sizes(z + 1) Then sizes(z + 1) = Len(flex.Text)
   Next
   rs.MoveNext
Loop

PBar.Value = 0
PBar.Visible = False
Command2.Enabled = False

flex.ColWidth(0) = 0
frmView!flex2.ColWidth(0) = 0
For z = 1 To smax
   flex.ColWidth(z) = sizes(z) * 100
   frmView!flex2.ColWidth(z) = sizes(z) * 100
Next
flex.Redraw = True
flex.Visible = True
conn.Close

End Sub

Private Sub indxcr_Click()

If TableName = "" Then
   MsgBox "You must first select a Table."
   Exit Sub
End If

frmCreateIndex.Show 1
List2_Click

End Sub

Private Sub indxdr_Click()

If TableName = "" Then
   MsgBox "You must first select a Table."
   Exit Sub
End If

frmDropIndex.Show 1
List2_Click

End Sub

Private Sub List1_Click()

frmMain.MousePointer = 11
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
For z = 0 To List1.ListCount - 1
  If List1.Selected(z) Then
     DBName = List1.List(z)
     Exit For
  End If
Next
frmMain.Caption = "SQL Developer (Server " + Server + ") Database " + DBName

List2.Clear
List3.Clear
List4.Clear
List5.Clear
keystr = ""
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

Set rs = conn.OpenSchema(adSchemaTables)
Do While Not rs.EOF
    If chkSys.Value = 0 Then
       If UCase(Left(rs!table_name, 3)) <> "SYS" Then List2.AddItem rs!table_name
    Else
       List2.AddItem rs!table_name
    End If
    rs.MoveNext
Loop

cmd1.CommandText = "sp_help"
Set rs = cmd1.Execute
rs.Filter = "Object_type='stored procedure'"
List4.Clear
keystr = ""
Do While Not rs.EOF
    If chkDT.Value = 0 Then
       If Left(rs!Name, 3) <> "dt_" Then List4.AddItem rs!Name
    Else
       List4.AddItem rs!Name
    End If
    rs.MoveNext
Loop
    
cmd1.CommandText = "sp_help"
Set rs = cmd1.Execute
rs.Filter = "Object_type='trigger'"
List5.Clear
Do While Not rs.EOF
    List5.AddItem rs!Name
    rs.MoveNext
Loop
conn.Close
TableName = ""
frmMain.MousePointer = 1

End Sub

Private Sub List2_Click()

frmMain.MousePointer = 11
List3.Clear
flex.Visible = False
On Error Resume Next
oldtext = Text1.Text

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
For z = 0 To List2.ListCount - 1
  If List2.Selected(z) Then
     TableName = List2.List(z)
     Exit For
  End If
Next
frmMain.Caption = "SQL Developer (Server " + Server + ") Database " + DBName + " | Table " + TableName

If InStr(TableName, " ") Then
   Text1.Text = "Select * from [" + TableName + "]"
Else
   Text1.Text = "Select * from " + TableName
End If

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

    Sql = "sp_columns [" + TableName + "]"
    cmd1.CommandText = Sql
    Set rs = cmd1.Execute
    cntset = 1
    Do While Not rs.EOF
        DStr = rs!column_name + ":" + rs!type_name + ":" + Str(rs!Length)
        List3.AddItem DStr
        If cntset Then
           cntset = 0
           CountSql = "Select count([" + rs!column_name + "]) as cmax from " + TableName
        End If
        rs.MoveNext
    Loop

Dim keys(500) As String
kmax = 0

   Sql = "sp_helpindex " + TableName
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.Fields.Count = 0 Then GoTo DoNext
   Do While Not rs.EOF
      pos = -1
      For z = 0 To List3.ListCount - 1
         clname$ = LCase(Left(List3.List(z), InStr(List3.List(z), ":") - 1))
         If InStr(LCase(rs!index_keys), clname$) Then
             pos = z
          kd$ = LCase(rs!index_description)
          If InStr(kd$, "primary key") Then
             ks$ = "PK"
          Else
             ks$ = "Indx"
          End If
          If InStr(kd$, "unique") Then ks$ = ks$ + "U"
          If InStr(kd$, "nonclust") Then ks$ = ks$ + "N"
          If InStr(kd$, "clustered") And InStr(kd$, "nonclus") = 0 Then ks$ = ks$ + "C"
          List3.List(pos) = List3.List(pos) + " (" + ks$ + ")"
         End If
       Next
       If InStr(rs!index_keys, ",") Then
          kmax = kmax + 1
          keys(kmax) = rs!index_keys
       End If
       rs.MoveNext
   Loop

   If kmax Then
      List3.AddItem "========Key Info =========="
      For z = 1 To kmax
         List3.AddItem keys(z)
      Next
   End If
   
DoNext:
conn.Close
Call ClearCombos(3)
Call ClearCombos(4)
Call ClearCombos(5)
On Error GoTo 0
frmMain.MousePointer = 1

End Sub

Private Sub List2_DblClick()

Text1.Text = oldtext
TableOrCol = "Table"
frmViewCols.Show 1
If TableOrCol <> "Table" Then
   For z = 0 To List2.ListCount - 1
      If List2.Selected(z) Then List2.Selected(z) = False
      If List2.List(z) = TableOrCol Then List2.Selected(z) = True
   Next
   List2_Click
End If

End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then
   If Len(keystr) < 1 Then
      keystr = ""
      Exit Sub
   Else
      keystr = Left(keystr, Len(keystr) - 1)
   End If
Else
   keystr = keystr + Chr(KeyAscii)
End If

For z = 0 To List2.ListCount - 1
   If Left(LCase(List2.List(z)), Len(keystr)) = LCase(keystr) Then
      List2.TopIndex = z
      Exit For
   End If
Next

End Sub

Private Sub List2_LostFocus()

keystr = ""

End Sub

Private Sub List3_Click()

Call ClearCombos(4)
Call ClearCombos(5)
oldtext = Text1.Text

For z = 0 To List3.ListCount - 1
   If List3.Selected(z) Then
      pos = z
      Exit For
   End If
Next

colname = Left(List3.List(pos), InStr(List3.List(pos), ":") - 1)
Text1.Text = Text1.Text + " " + colname + " "
Text1.SetFocus

End Sub

Private Sub List3_DblClick()

Text1.Text = oldtext
TableOrCol = "Col"
frmViewCols.Show 1

End Sub

Private Sub List4_Click()

For z = 0 To List4.ListCount - 1
   If List4.Selected(z) Then
      indxc = z
      Exit For
   End If
Next

Call ClearCombos(3)
Call ClearCombos(5)
If indxc Then List4.Selected(indxc) = True

End Sub

Private Sub List4_DblClick()

frmMain.MousePointer = 11
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
For z = 0 To List4.ListCount - 1
  If List4.Selected(z) Then
     SPName = List4.List(z)
     Exit For
  End If
Next

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

cmd1.CommandText = "sp_helptext " + SPName
Set rs = cmd1.Execute
strText = ""
Do While Not rs.EOF
    strText = strText + rs!Text
    rs.MoveNext
Loop
Text1.Text = strText
frmMain.MousePointer = 1
frmText.Show 1
    
End Sub

Private Sub List4_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then
   If Len(keystr) < 1 Then
      keystr = ""
      Exit Sub
   Else
      keystr = Left(keystr, Len(keystr) - 1)
   End If
Else
   keystr = keystr + Chr(KeyAscii)
End If

For z = 0 To List4.ListCount - 1
   If Left(LCase(List4.List(z)), Len(keystr)) = LCase(keystr) Then
      List4.TopIndex = z
      Exit For
   End If
Next

End Sub

Private Sub List4_LostFocus()

keystr = ""

End Sub

Private Sub List5_DblClick()

frmMain.MousePointer = 11
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
For z = 0 To List5.ListCount - 1
  If List5.Selected(z) Then
     TRName = List5.List(z)
     Exit For
  End If
Next

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

cmd1.CommandText = "sp_helptext " + TRName
Set rs = cmd1.Execute
strText = ""
Do While Not rs.EOF
    strText = strText + rs!Text
    rs.MoveNext
Loop
Text1.Text = strText
frmMain.MousePointer = 1
frmText.Show 1
conn.Close

End Sub

Private Sub List5_Click()

For z = 0 To List5.ListCount - 1
   If List5.Selected(z) Then
      indxc = z
      Exit For
   End If
Next

Call ClearCombos(3)
Call ClearCombos(4)
If indxc Then List5.Selected(indxc) = True

End Sub

Private Sub opt_Click()

frmOptions.Show 1
Form_Load

End Sub

Private Sub tbalt_Click()

If TableName = "" Then
   MsgBox "You must first select a table."
   Exit Sub
End If

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

ASql = "Alter Table dbo." + TableName + " (" + vbCrLf
Sql = "sp_columns [" + TableName + "]"
cmd1.CommandText = Sql
Set rs = cmd1.Execute
Do While Not rs.EOF
    colname$ = "[" + rs!column_name + "]"
    ASql = ASql + colname$ + " [" + rs!type_name + "] "
    Select Case rs!type_name
       Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
           ASql = ASql + "(" + Trim(Str(rs!Length)) + ")"
    End Select
    If rs!Is_Nullable = "YES" Then
       ASql = ASql + " NULL," + vbCrLf
    Else
       ASql = ASql + " NOT NULL," + vbCrLf
    End If
    rs.MoveNext
Loop
ASql = Left(ASql, Len(ASql) - 2) + vbCrLf + ") ON [PRIMARY]"
Text1.Text = ASql
frmText!Text1.Text = Text1.Text
frmText.Show
conn.Close

End Sub

Private Sub popdel_Click()

flex.Row = 0
flex.Col = CurrCol
colname = flex.Text
sysbut% = MsgBox("Are you sure you want to delete the " + colname + " column?", 4, "Delete Confirm")
If sysbut% <> vbYes Then Exit Sub
frmMain.MousePointer = 11
flex.Redraw = False
For z = CurrCol To flex.Cols - 2
  For zz = 0 To flex.Rows - 1
    flex.Row = zz
    flex.Col = z + 1
    d$ = flex.Text
    flex.Text = ""
    flex.Col = z
    flex.Text = d$
  Next
Next
For z = CurrCol To flex.Cols - 2
   cw = flex.ColWidth(z + 1)
   flex.Col = z
   flex.ColWidth(z) = cw
Next

flex.Cols = flex.Cols - 1
flex.Redraw = True
flex.Refresh
frmMain.MousePointer = 1

End Sub

Private Sub popedt_Click()

popmenu.Visible = False
EditOrInsert = "Edit"
frmEditInsertRow.Show 1

End Sub

Private Sub popext_Click()

popmenu.Visible = False

End Sub

Private Sub popins_Click()

popmenu.Visible = False
EditOrInsert = "Insert"
frmEditInsertRow.Show 1

End Sub

Private Sub scrdb_Click()

If DBName = "" Then
   MsgBox "You must select a Database."
   Exit Sub
End If

ScriptType = "DB"
frmScript.Show 1

End Sub

Private Sub scrproc_Click()

ScriptType = "Proc"
frmScript.Show 1

End Sub

Private Sub scrtb_Click()

If TableName = "" Then
   MsgBox "You must select a Table."
   Exit Sub
End If

ScriptType = "Table"
frmScript.Show 1

End Sub

Private Sub ser_Click()

frmSearchMain.Show '1

End Sub

Private Sub spcret_Click()

frmCreateSP.Show 1
List1_Click

End Sub

Private Sub spdebug_Click()

If DBName = "" Then
   MsgBox "You must select a database first."
   Exit Sub
End If

If List4.ListCount = 0 Then
   MsgBox "There are no stored procedures at this time."
   Exit Sub
End If

frmDebug.Show 1

End Sub

Private Sub tbclr_Click()

If TableName = "" Then
   MsgBox "You must first select a table."
   Exit Sub
End If

sysbut% = MsgBox("Are you sure you want to clear the table " + TableName + "?", 4, "Clear Confirm")
If sysbut% <> vbYes Then Exit Sub

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Sql = "Delete from " + TableName
cmd1.CommandText = Sql
cmd1.Execute
conn.Close

End Sub

Private Sub tbcreate_Click()

If DBName = "" Then
   MsgBox "You must select a database first."
   Exit Sub
End If

frmTableCreate.Show 1
List1_Click

End Sub

Private Sub tbdep_Click()

If DBName = "" Then
   MsgBox "A Database must first be selected."
   Exit Sub
End If

frmMain.MousePointer = 11
Text1.Text = ""
d$ = ""
numdep& = 0
On Error GoTo errorfound

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
sp = 0
ep = List2.ListCount - 1
For z = sp To ep
   If List2.Selected(z) Then
      sp = z
      ep = z
      Exit For
   End If
Next

For z = sp To ep
   d$ = d$ + "Table " + List2.List(z) + vbCrLf
   Sql = "sp_depends " + List2.List(z)
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
      d$ = d$ + Chr(9) + "No Dependencies." + vbCrLf
   Else
      Do While Not rs.EOF
         d$ = d$ + Chr(9) + rs!Name + " - " + rs!Type + vbCrLf
         numdep& = numdep& + 1
         rs.MoveNext
      Loop
   End If
Next
conn.Close

On Error GoTo 0
Text1.Text = d$
frmMain.MousePointer = 1
frmText.Show 1
Text1.Text = ""
If numdep& Then
   If numdep& = 1 Then
      lbStat.Caption = "One Dependency returned."
      lbStat.Refresh
   Else
      lbStat.Caption = Format$(numdep&, "###,###,##0") + " Dependencies returned."
      lbStat.Refresh
   End If
Else
   lbStat.Caption = "No Dependencies were returned."
   lbStat.Refresh
End If
Exit Sub

errorfound:
   Debug.Print Err.Number, Sql, Err.Description
   Resume Next

End Sub

Private Sub Text1_Change()

Command2.Enabled = False
If Text1.Text = "" Then
   Command1.Enabled = False
   Command3.Enabled = False
Else
   Command1.Enabled = True
   Command3.Enabled = True
End If

End Sub

Private Sub Text1_DblClick()

frmText!Text1.Text = Text1.Text
frmText.Show 1

End Sub

Public Sub ClearCombos(box As Integer)

Select Case box
   Case 3
      For z = 0 To List3.ListCount - 1
         If List3.Selected(z) Then
            List3.Selected(z) = False
            Exit For
         End If
      Next
   Case 4
      For z = 0 To List4.ListCount - 1
         If List4.Selected(z) Then
            List4.Selected(z) = False
            Exit For
         End If
      Next
   Case 5
      For z = 0 To List5.ListCount - 1
         If List5.Selected(z) Then
            List5.Selected(z) = False
            Exit For
         End If
      Next
End Select

End Sub

Public Function DetermineIfNotSQLStatement(s$) As Long

DetermineIfNotSQLStatement = True
s$ = LCase(s$)
If Left(s$, 7) = "select " Or Left(s$, 7) = "update " Or Left(s$, 5) = "drop " Or Left(s$, 7) = "create " Or Left(s$, 7) = "insert " Then
   DetermineIfNotSQLStatement = False
End If

End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
   Case 116
       Command1_Click
End Select

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button <> 2 Then Exit Sub
Combo1.Visible = True

End Sub

Private Sub Timer1_Timer()

If List2.ListCount = 0 Then
   List2.Visible = False
   Label2.Visible = False
   chkSys.Visible = False
Else
   List2.Visible = True
   Label2.Visible = True
   chkSys.Visible = True
End If

If List3.ListCount = 0 Then
   List3.Visible = False
   Label3.Visible = False
Else
   List3.Visible = True
   Label3.Visible = True
End If

If List4.ListCount = 0 Then
   List4.Visible = False
   Label4.Visible = False
   chkDT.Visible = False
Else
   List4.Visible = True
   Label4.Visible = True
   chkDT.Visible = True
End If

If List5.ListCount = 0 Then
   List5.Visible = False
   Label5.Visible = False
Else
   List5.Visible = True
   Label5.Visible = True
End If

Timer1.Interval = 500

End Sub

Private Sub utldel_Click()

If DBName = "" Then
   MsgBox "The Database needs to be selected.  If columns are going to be deleted, the table needs to be selected first."
   Exit Sub
End If

frmDelete.Show 1
If ReloadedNeeded Then
   ReloadedNeeded = False
   List1_Click
End If

If TableName <> "" Then
   List2_Click
End If
   
End Sub

Private Sub utlpnt_Click()

frmMain.PrintForm

End Sub

Private Sub utlren_Click()

If DBName = "" Then
   MsgBox "The Database needs to be selected.  If columns are going to be renamed, the table needs to be selected first."
   Exit Sub
End If

frmRename.Show 1

End Sub

Private Sub utlsp_Click()

If List4.ListCount = 0 Then
   MsgBox "No Stored Procedures at this time."
   Exit Sub
End If

frmMoveStoredProcs.Show 1

End Sub

Private Sub vwdb_Click()

Dim vdat(50000, 5)
vmax = 0

On Error Resume Next
frmMain.MousePointer = 11
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
Set tcmd1 = New ADODB.Command
For v = 0 To List1.ListCount - 1
   lbStat.Caption = "Scanning " + List1.List(v)
   lbStat.Refresh
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + List1.List(v) + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   tcmd1.ActiveConnection = conn
   Set rs = conn.OpenSchema(adSchemaTables)
   Do While Not rs.EOF
       If UCase(Left(rs!table_name, 3)) <> "SYS" Then
           Sql = "sp_helpindex " + rs!table_name
           tcmd1.CommandText = Sql
           Set trs = tcmd1.Execute
           If trs.Fields.Count = 0 Then GoTo DoNext
           Do While Not trs.EOF
              vmax = vmax + 1
              vdat(vmax, 1) = List1.List(v)
              vdat(vmax, 2) = rs!table_name
              vdat(vmax, 3) = trs!index_keys
              vdat(vmax, 4) = trs!index_name
              vdat(vmax, 5) = trs!index_description
              trs.MoveNext
           Loop
       End If
DoNext:
       rs.MoveNext
   Loop
conn.Close
Next
On Error GoTo 0
flex.Redraw = False
flex.Visible = False
flex.Rows = vmax + 1
flex.Cols = 6
flex.Row = 0
flex.Col = 1
flex.Text = "Database"
flex.Col = 2
flex.Text = "Table"
flex.Col = 3
flex.Text = "Column"
flex.Col = 4
flex.Text = "Index Name"
flex.Col = 5
flex.Text = "Index Description"

frmView!flex2.Rows = vmax + 1
frmView!flex2.Cols = 6
frmView!flex2.Row = 0
frmView!flex2.Col = 1
frmView!flex2.Text = "DataBase"
frmView!flex2.Col = 2
frmView!flex2.Text = "Table"
frmView!flex2.Col = 3
frmView!flex2.Text = "Column"
frmView!flex2.Col = 4
frmView!flex2.Text = "Index Name"
frmView!flex2.Col = 5
frmView!flex2.Text = "Index Description"

For z = 1 To vmax
   flex.Row = z
   frmView!flex2.Row = z
   For zz = 1 To 5
      flex.Col = zz
      flex.Text = vdat(z, zz)
      frmView!flex2.Col = zz
      frmView!flex2.Text = vdat(z, zz)
   Next
Next

flex.ColWidth(0) = 0
flex.ColWidth(1) = 2000
flex.ColWidth(2) = 2000
flex.ColWidth(3) = 2000
flex.ColWidth(4) = 3000
flex.ColWidth(5) = 8000

frmView!flex2.ColWidth(0) = 0
frmView!flex2.ColWidth(1) = 2000
frmView!flex2.ColWidth(2) = 2000
frmView!flex2.ColWidth(3) = 2000
frmView!flex2.ColWidth(4) = 3000
frmView!flex2.ColWidth(5) = 8000

flex.Redraw = True
flex.Visible = True
frmMain.MousePointer = 1
lbStat.Caption = "Ready..."
lbStat.Refresh
conn.Close

End Sub

Private Sub vwdbc_Click()

If DBName = "" Then
   MsgBox "You must first select a Database."
   Exit Sub
End If

Dim vdat(50000, 4) As String
vmax = 0

On Error Resume Next
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
Set tcmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
tcmd1.ActiveConnection = conn
Set rs = conn.OpenSchema(adSchemaTables)
Do While Not rs.EOF
    If UCase(Left(rs!table_name, 3)) <> "SYS" Then
        Sql = "sp_helpindex " + rs!table_name
        tcmd1.CommandText = Sql
        Set trs = tcmd1.Execute
        If trs.Fields.Count = 0 Then GoTo DoNext
        Do While Not trs.EOF
           vmax = vmax + 1
           vdat(vmax, 1) = rs!table_name
           vdat(vmax, 2) = trs!index_keys
           vdat(vmax, 3) = trs!index_name
           vdat(vmax, 4) = trs!index_description
           trs.MoveNext
       Loop
    End If
DoNext:
rs.MoveNext
Loop

conn.Close
On Error GoTo 0
flex.Redraw = False
flex.Visible = False
flex.Rows = vmax + 1
flex.Cols = 5
flex.Row = 0
flex.Col = 1
flex.Text = "Table"
flex.Col = 2
flex.Text = "Column"
flex.Col = 3
flex.Text = "Index Name"
flex.Col = 4
flex.Text = "Index Description"

frmView!flex2.Rows = vmax + 1
frmView!flex2.Cols = 5
frmView!flex2.Row = 0
frmView!flex2.Col = 1
frmView!flex2.Text = "Table"
frmView!flex2.Col = 2
frmView!flex2.Text = "Column"
frmView!flex2.Col = 3
frmView!flex2.Text = "Index Name"
frmView!flex2.Col = 4
frmView!flex2.Text = "Index Description"

For z = 1 To vmax
   flex.Row = z
   frmView!flex2.Row = z
   For zz = 1 To 4
      flex.Col = zz
      flex.Text = vdat(z, zz)
      frmView!flex2.Col = zz
      frmView!flex2.Text = vdat(z, zz)
   Next
Next

flex.ColWidth(0) = 0
flex.ColWidth(1) = 2000
flex.ColWidth(2) = 2000
flex.ColWidth(3) = 3000
flex.ColWidth(4) = 8000

frmView!flex2.ColWidth(0) = 0
frmView!flex2.ColWidth(1) = 2000
frmView!flex2.ColWidth(2) = 2000
frmView!flex2.ColWidth(3) = 3000
frmView!flex2.ColWidth(4) = 8000

flex.Redraw = True
flex.Visible = True

End Sub
