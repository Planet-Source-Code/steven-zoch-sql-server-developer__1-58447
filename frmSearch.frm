VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Searching"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Convert to Grid"
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   8775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   3960
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Search"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbStat 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StopSearch As Boolean
Public parm1 As String
Public parm2 As String
Public tmax As Long
Public numfound As Long
Public SearchDone As Boolean

Private Sub Command1_Click()

StopSearch = True

End Sub

Private Sub Command2_Click()

SearchDone = False
SearchInProgress = False
Unload frmSearch
Unload frmSearchShow

End Sub

Private Sub Command3_Click()

frmView.Show 1

End Sub

Private Sub Form_Activate()

Command3.Visible = False
If SearchDone Then Exit Sub
SearchDone = True
pos = InStr(SearchType, " ")
parm1 = Left(SearchType, pos - 1)
parm2 = Mid(SearchType, pos + 1)
numfound = 0
SearchInProgress = True

Select Case parm1
   Case "default", "table"
      Select Case parm2
         Case "col"
            Call DefaultCol
         Case "dt"
            Call DefaultDataType
         Case "dsc"
            Call DefaultDataCol
         Case "dsd"
            Call DefaultDataData
            Command3.Visible = True
         Case "dss"
            Call DefaultDataStored
         Case "dst"
            Call DefaultDataTrigger
      End Select
   Case "entire"
      Select Case parm2
         Case "col"
            Call EntireCol
         Case "dt"
            Call EntireDataType
         Case "dsc"
            Call EntireDataCol
         Case "dsd"
            Call EntireDataData
            Command3.Visible = True
         Case "dss"
            Call EntireDataStored
         Case "dst"
            Call EntireDataTrigger
      End Select
End Select

lbStat.Caption = ""
lbStat.Refresh
If numfound = 0 Then List1.AddItem "No Search Results..."

End Sub

Private Sub Form_Load()

List1.Clear
List2.Clear
keystr = ""
StopSearch = False
Command2.Enabled = False
lbStat.Caption = ""
PBar.Visible = False
Text2.Visible = False

End Sub

Public Sub DefaultCol()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 0 To frmMain!List2.ListCount - 1
   curtab = frmMain!List2.List(z)
   If parm1 = "default" Or curtab = TableName Then
      lbStat.Caption = "Searching DB " + DBName + " Table-" + curtab
      Sql = "sp_columns [" + curtab + "]"
      cmd1.CommandText = Sql
      Set rs = cmd1.Execute
      Do While Not rs.EOF
          If LCase(SearchStr) = LCase(rs!column_name) Then
             Call FoundOne
             List1.AddItem "DB-" + DBName + " Table-" + curtab
          End If
          DoEvents
          If StopSearch Then
              Exit For
          End If
          rs.MoveNext
      Loop
   End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub DefaultDataType()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 0 To frmMain!List2.ListCount - 1
   curtab = frmMain!List2.List(z)
   If parm1 = "default" Or curtab = TableName Then
       lbStat.Caption = "Searching DB " + DBName + " Table-" + curtab
       Sql = "sp_columns [" + curtab + "]"
       cmd1.CommandText = Sql
       Set rs = cmd1.Execute
       Do While Not rs.EOF
           If LCase(SearchStr) = LCase(rs!type_name) Then
              Call FoundOne
              List1.AddItem "DB-" + DBName + " Table-" + curtab + " Column-" + rs!column_name
           End If
           DoEvents
           If StopSearch Then
               Exit For
           End If
           rs.MoveNext
       Loop
    End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub DefaultDataCol()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 0 To frmMain!List2.ListCount - 1
   curtab = frmMain!List2.List(z)
   If parm1 = "default" Or curtab = TableName Then
       lbStat.Caption = "Searching DB " + DBName + " Table-" + curtab
       Sql = "sp_columns [" + curtab + "]"
       cmd1.CommandText = Sql
       Set rs = cmd1.Execute
       Do While Not rs.EOF
           If InStr(LCase(rs!column_name), LCase(SearchStr)) Then
              Call FoundOne
              List1.AddItem "DB-" + DBName + " Table-" + curtab + " Column-" + rs!column_name
           End If
           DoEvents
           If StopSearch Then
               Exit For
           End If
           rs.MoveNext
       Loop
    End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub DefaultDataData()

Dim DStr As String

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
For z = 0 To frmMain!List2.ListCount - 1
   curtab = frmMain!List2.List(z)
   If parm1 = "default" Or curtab = TableName Then
      lbStat.Caption = "Searching DB " + DBName + " Table-" + curtab
      Sql = "select * from [" + curtab + "]"
       cmd1.CommandText = Sql
       Set rs = cmd1.Execute
       Do While Not rs.EOF
           For zz = 0 To rs.Fields.Count - 1
              Select Case rs.Fields(zz).Type
                 Case 129, 130, 202, 200, 135, 205, 203, 201, 72
                    If IsNull(rs.Fields(zz).Value) Then
                       d$ = "''"
                    Else
                       If rs.Fields(zz).Type = 135 Then
                          d$ = Str(rs.Fields(zz).Value)
                       Else
                          d$ = rs.Fields(zz).Value
                       End If
                    End If
                 Case 204, 128
                    acnv = StrConv(trs.Fields(zz).Value, vbUnicode)
                    If IsNull(acnv) Then
                        hxstr = ""
                    Else
                        hxstr = "0x"
                        For y = 1 To Len(acnv)
                            hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
                        Next
                    End If
                    d$ = Trim(hxstr)
                 Case 20, 131, 5, 3, 6, 2, 17, 11, 4
                    If IsNull(rs.Fields(zz).Value) Then
                       d$ = "0"
                    Else
                       d$ = Trim(Str(rs.Fields(zz).Value))
                    End If
              End Select
              If InStr(LCase(d$), LCase(SearchStr)) Then
                 Call FoundOne
                 List1.AddItem "DB-" + DBName + " Table-" + curtab + " Column-" + rs.Fields(zz).Name + " Data=" + d$
                 DStr = ""
                 For ds = 0 To rs.Fields.Count - 1
                    DStr = DStr + rs.Fields(ds).Name + "|"
                    Select Case rs.Fields(ds).Type
                       Case 129, 130, 202, 200, 135, 205, 203, 201, 72
                          If IsNull(rs.Fields(ds).Value) Then
                             d$ = "''"
                          Else
                             If rs.Fields(ds).Type = 135 Then
                                d$ = Str(rs.Fields(ds).Value)
                             Else
                                d$ = rs.Fields(ds).Value
                             End If
                          End If
                       Case 204, 128
                          acnv = StrConv(trs.Fields(ds).Value, vbUnicode)
                          If IsNull(acnv) Then
                              hxstr = ""
                          Else
                              hxstr = "0x"
                              For y = 1 To Len(acnv)
                                  hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
                              Next
                          End If
                          d$ = Trim(hxstr)
                       Case 20, 131, 5, 3, 6, 2, 17, 11, 4
                          If IsNull(rs.Fields(ds).Value) Then
                             d$ = "0"
                          Else
                             d$ = Trim(Str(rs.Fields(ds).Value))
                          End If
                    End Select
                    DStr = DStr + d$ + "|"
                 Next
                    List2.AddItem DStr
                 End If
                 DoEvents
                 If StopSearch Then
                     Exit For
                 End If
              Next
              rs.MoveNext
          Loop
    End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub DefaultDataStored()

Dim strText As String
Dim linenum As Long
Dim instrval As Long

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

For z = 0 To frmMain!List4.ListCount - 1
   curtab = frmMain!List4.List(z)
   If parm1 = "default" Or curtab = TableName Then
      lbStat.Caption = "Searching DB " + DBName + " Proc-" + curtab
      cmd1.CommandText = "sp_helptext " + curtab
      Set rs = cmd1.Execute
      strText = ""
      Do While Not rs.EOF
          strText = strText + rs!Text
          rs.MoveNext
      Loop
      If InStr(LCase(strText), LCase(SearchStr)) Then
           instrval = InStr(LCase(strText), LCase(SearchStr))
           Call FoundOne
           List1.AddItem "DB-" + DBName + " Procedure-" + curtab
           List2.AddItem strText
      End If
      DoEvents
      If StopSearch Then
          Exit For
      End If
   End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub DefaultDataTrigger()

Dim strText As String
Dim linenum As Long
Dim instrval As Long

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn

For z = 0 To frmMain!List5.ListCount - 1
   curtab = frmMain!List5.List(z)
   If parm1 = "default" Or curtab = TableName Then
      lbStat.Caption = "Searching DB " + DBName + " Proc-" + curtab
      cmd1.CommandText = "sp_helptext " + curtab
      Set rs = cmd1.Execute
      strText = ""
      linenum = 0
      Do While Not rs.EOF
          linenum = linenum + 1
          strText = strText + rs!Text
          rs.MoveNext
      Loop
      If InStr(LCase(strText), LCase(SearchStr)) Then
              'Found
              instrval = InStr(LCase(strText), LCase(SearchStr))
              Call FoundOne
              List1.AddItem "DB-" + DBName + " Trigger-" + curtab
              List2.AddItem strText
      End If
      DoEvents
      If StopSearch Then
          Exit For
      End If
   End If
Next
Command2.Enabled = True
conn.Close

End Sub
Public Sub EntireCol()

Call LoadArray("Table")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
PBar.Value = 0
PBar.Visible = True
For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   curtab = ToDo(z, 2)
   lbStat.Caption = "Searching DB " + curDBName + " Table-" + curtab
   Sql = "sp_columns [" + curtab + "]"
    cmd1.CommandText = Sql
    Set rs = cmd1.Execute
    Do While Not rs.EOF
        If LCase(SearchStr) = LCase(rs!column_name) Then
           Call FoundOne
           List1.AddItem "DB-" + DBName + " Table-" + curtab
        End If
        DoEvents
        If StopSearch Then
            Exit For
        End If
        rs.MoveNext
    Loop
    conn.Close
Next
Command2.Enabled = True
PBar.Visible = False

End Sub
Public Sub EntireDataType()

Call LoadArray("Table")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
PBar.Value = 0
PBar.Visible = True
For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   curtab = ToDo(z, 2)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   lbStat.Caption = "Searching DB " + curDBName + " Table-" + curtab
   Sql = "sp_columns [" + curtab + "]"
    cmd1.CommandText = Sql
    Set rs = cmd1.Execute
    Do While Not rs.EOF
        If LCase(SearchStr) = LCase(rs!type_name) Then
           Call FoundOne
           List1.AddItem "DB-" + curDBName + " Table-" + curtab + " Column-" + rs!column_name
        End If
        DoEvents
        If StopSearch Then
            Exit For
        End If
        rs.MoveNext
    Loop
    conn.Close
Next
Command2.Enabled = True
PBar.Visible = False

End Sub
Public Sub EntireDataCol()

Call LoadArray("Table")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
On Error Resume Next
PBar.Value = 0
PBar.Visible = True
For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   curtab = ToDo(z, 2)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   lbStat.Caption = "Searching DB " + curDBName + " Table-" + curtab
   Sql = "sp_columns [" + curtab + "]"
    cmd1.CommandText = Sql
    Set rs = cmd1.Execute
    Do While Not rs.EOF
        If InStr(LCase(rs!column_name), LCase(SearchStr)) Then
           Call FoundOne
           List1.AddItem "DB-" + curDBName + " Table-" + curtab + " Column-" + rs!column_name
        End If
        DoEvents
        If StopSearch Then
            Exit For
        End If
        rs.MoveNext
    Loop
    conn.Close
Next
Command2.Enabled = True
PBar.Visible = False

End Sub
Public Sub EntireDataData()

Call LoadArray("Table")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
On Error Resume Next
PBar.Value = 0
PBar.Visible = True
For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   curtab = ToDo(z, 2)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   lbStat.Caption = "Searching DB " + curDBName + " Table-" + curtab
   Sql = "select * from [" + curtab + "]"
    cmd1.CommandText = Sql
    Set rs = cmd1.Execute
    Do While Not rs.EOF
        For zz = 0 To rs.Fields.Count - 1
           Select Case rs.Fields(zz).Type
              Case 129, 130, 202, 200, 135, 205, 203, 201, 72
                 If IsNull(rs.Fields(zz).Value) Then
                    d$ = "''"
                 Else
                    If rs.Fields(zz).Type = 135 Then
                       d$ = Str(rs.Fields(zz).Value)
                    Else
                       d$ = rs.Fields(zz).Value
                    End If
                 End If
              Case 204, 128
                 acnv = StrConv(trs.Fields(zz).Value, vbUnicode)
                 If IsNull(acnv) Then
                     hxstr = ""
                 Else
                     hxstr = "0x"
                     For y = 1 To Len(acnv)
                         hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
                     Next
                 End If
                 d$ = Trim(hxstr)
              Case 20, 131, 5, 3, 6, 2, 17, 11, 4
                 If IsNull(rs.Fields(zz).Value) Then
                    d$ = "0"
                 Else
                    d$ = Trim(Str(rs.Fields(zz).Value))
                 End If
           End Select
           If InStr(LCase(d$), LCase(SearchStr)) Then
              Call FoundOne
              List1.AddItem "DB-" + curDBName + " Table-" + curtab + " Column-" + rs.Fields(z).Name + " Data=" + d$
           End If
           DoEvents
           If StopSearch Then
               Exit For
           End If
        Next
        rs.MoveNext
    Loop
    conn.Close
Next
Command2.Enabled = True
PBar.Visible = False
On Error GoTo 0

End Sub
Public Sub EntireDataStored()

Dim strText As String
Dim linenum As Long
Dim instrval As Long

Call LoadArray("Proc")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
On Error Resume Next
PBar.Value = 0
PBar.Visible = True

For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   curtab = ToDo(z, 2)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   lbStat.Caption = "Searching DB " + curDBName + " Proc-" + curtab
   cmd1.CommandText = "sp_helptext " + curtab
   Set rs = cmd1.Execute
   strText = ""
   linenum = 0
   Do While Not rs.EOF
       strText = strText + rs!Text
       If IsNull(rs!Text) Then Exit Do
       rs.MoveNext
   Loop
   If InStr(LCase(strText), LCase(SearchStr)) Then
           instrval = InStr(LCase(strText), LCase(SearchStr))
           Call FoundOne
           List1.AddItem "DB-" + curDBName + " Procedure-" + curtab
           List2.AddItem strText
   End If
   DoEvents
   If StopSearch Then
       Exit For
   End If
   conn.Close
Next
Command2.Enabled = True
On Error GoTo 0
PBar.Visible = False

End Sub
Public Sub EntireDataTrigger()

Dim strText As String
Dim linenum As Long
Dim instrval As Long

Call LoadArray("Trigger")
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
On Error Resume Next
PBar.Value = 0
PBar.Visible = True

For z = 1 To tmax
   PBar.Value = Int((z / tmax) * 100)
   PBar.Refresh
   curDBName = ToDo(z, 1)
   curtab = ToDo(z, 2)
   strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + curDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   lbStat.Caption = "Searching DB " + curDBName + " Trigger-" + curtab
   cmd1.CommandText = "sp_helptext " + curtab
   Set rs = cmd1.Execute
   strText = ""
   linenum = 0
   Do While Not rs.EOF
       linenum = linenum + 1
       strText = strText + rs!Text
       If IsNull(rs!Text) Then Exit Do
       rs.MoveNext
   Loop
   If InStr(LCase(strText), LCase(SearchStr)) Then
           instrval = InStr(LCase(strText), LCase(SearchStr))
           Call FoundOne
           List1.AddItem "DB-" + curDBName + " Trigger-" + curtab
           List2.AddItem strText
   End If
   DoEvents
   If StopSearch Then
       Exit For
   End If
   conn.Close
Next
Command2.Enabled = True
On Error GoTo 0
PBar.Visible = False

End Sub

Public Sub LoadArray(t$)

Erase ToDo
tmax = 0
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
Select Case t$
   Case "Table"
      For z = 0 To frmMain!List1.ListCount - 1
         strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + frmMain!List1.List(z) + ";uid=" + UID + ";pwd=" + PWD + ";"
         conn.Open strConnect
         cmd1.ActiveConnection = conn
         Set rs = conn.OpenSchema(adSchemaTables)
         Do While Not rs.EOF
            If UCase(Left(rs!table_name, 3)) <> "SYS" Then
                tmax = tmax + 1
                ToDo(tmax, 1) = frmMain!List1.List(z)
                ToDo(tmax, 2) = rs!table_name
             End If
             rs.MoveNext
         Loop
         conn.Close
       Next
   Case "Proc"
      For z = 0 To frmMain!List1.ListCount - 1
         strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + frmMain!List1.List(z) + ";uid=" + UID + ";pwd=" + PWD + ";"
         conn.Open strConnect
         cmd1.ActiveConnection = conn
         cmd1.CommandText = "sp_help"
         Set rs = cmd1.Execute
         rs.Filter = "Object_type='stored procedure'"
         Do While Not rs.EOF
              tmax = tmax + 1
              ToDo(tmax, 1) = frmMain!List1.List(z)
              ToDo(tmax, 2) = rs!Name
              rs.MoveNext
         Loop
         conn.Close
       Next
   Case "Trigger"
      For z = 0 To frmMain!List1.ListCount - 1
         strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + frmMain!List1.List(z) + ";uid=" + UID + ";pwd=" + PWD + ";"
         conn.Open strConnect
         cmd1.ActiveConnection = conn
         cmd1.CommandText = "sp_help"
         Set rs = cmd1.Execute
         rs.Filter = "Object_type='trigger'"
         Do While Not rs.EOF
              tmax = tmax + 1
              ToDo(tmax, 1) = frmMain!List1.List(z)
              ToDo(tmax, 2) = rs!Name
              rs.MoveNext
         Loop
         conn.Close
       Next
End Select

End Sub

Public Sub FoundOne()

numfound = numfound + 1
Text2.Text = Format$(numfound, "#,###,###,##0")
Text2.Refresh
Text2.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

SearchInProgress = False
SearchDone = False

End Sub

Private Sub List1_Click()

If parm2 = "dss" Or parm2 = "dst" Then
   Unload frmSearchShow
   For z = 0 To List1.ListCount - 1
      If List1.Selected(z) Then
          TextData = List2.List(z)
          Exit For
      End If
   Next
   frmSearchShow.Show
End If

End Sub

Private Sub List1_DblClick()

Dim cdat(5000, 2)
Unload frmSearchShow

If parm2 = "dss" Or parm2 = "dst" Then
   For z = 0 To List1.ListCount - 1
      If List1.Selected(z) Then
          frmText!Text1.Text = List2.List(z)
          Exit For
      End If
   Next
   frmText.Show 1
End If

If parm2 = "dsd" Then
   For z = 0 To List1.ListCount - 1
      If List1.Selected(z) Then
          t$ = List2.List(z)
          Exit For
      End If
   Next
   tstr = ""
   p = 0
   cm = 0
   For z = 1 To Len(t$)
      If Mid(t$, z, 1) = "|" Then
         p = p + 1
         If p = 3 Then p = 1
         If p = 1 Then cm = cm + 1
         cdat(cm, p) = ct$
         ct$ = ""
      Else
         ct$ = ct$ + Mid(t$, z, 1)
      End If
   Next
   ml = 0
   For z = 1 To cm
      If Len(cdat(z, 1)) > ml Then ml = Len(cdat(z, 1))
   Next
   For z = 1 To cm
      cdat(z, 1) = Left(cdat(z, 1) + Space$(ml), ml)
      cdat(z, 1) = Replace(cdat(z, 1), " ", ".")
      tstr = tstr + cdat(z, 1) + " " + cdat(z, 2) + vbCrLf
   Next
   frmText!Text1.Text = tstr
   frmText.Show 1
End If

End Sub
Public Sub DetermineString(d$)

           Select Case rs.Fields(zz).Type
              Case 129, 130, 202, 200, 135, 205, 203, 201, 72
                 If IsNull(rs.Fields(zz).Value) Then
                    d$ = "''"
                 Else
                    If rs.Fields(zz).Type = 135 Then
                       d$ = Str(rs.Fields(zz).Value)
                    Else
                       d$ = rs.Fields(zz).Value
                    End If
                 End If
              Case 204, 128
                 acnv = StrConv(trs.Fields(zz).Value, vbUnicode)
                 If IsNull(acnv) Then
                     hxstr = ""
                 Else
                     hxstr = "0x"
                     For y = 1 To Len(acnv)
                         hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
                     Next
                 End If
                 d$ = Trim(hxstr)
              Case 20, 131, 5, 3, 6, 2, 17, 11, 4
                 If IsNull(rs.Fields(zz).Value) Then
                    d$ = "0"
                 Else
                    d$ = Trim(Str(rs.Fields(zz).Value))
                 End If
           End Select

End Sub

