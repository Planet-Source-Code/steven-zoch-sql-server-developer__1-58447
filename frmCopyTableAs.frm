VERSION 5.00
Begin VB.Form frmCopyTableAs 
   Caption         =   "Copy Table As"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   Icon            =   "frmCopyTableAs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbServer 
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Text            =   "Combo2"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Structure and Data"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Structure Only"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Server"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "New Table"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Table to Copy"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmCopyTableAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then
   MsgBox "You must provide the new table name."
   Exit Sub
End If

Open "c:\temp.sql" For Output As #1

Set tconn = New ADODB.Connection
Set tcmd1 = New ADODB.Command
Set mconn = New ADODB.Connection
Set mcmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
tconn.Open strConnect
tcmd1.ActiveConnection = tconn

LogAction = "Copy Table As"
LogComment = "Table " + Combo1.Text + " copied as " + Text1.Text + " on server " + cmbServer.Text
tbname$ = Text1.Text
ASql = "Create Table dbo." + tbname$ + " ("
Sql = "sp_columns [" + Combo1.Text + "]"
tcmd1.CommandText = Sql
Set trs = tcmd1.Execute
Do While Not trs.EOF
    colname$ = "[" + trs!column_name + "]"
    ASql = ASql + colname$ + " [" + trs!type_name + "] "
    Select Case trs!type_name
       Case "binary", "char", "nchar", "nvarchar", "varbinary", "varchar"
           ASql = ASql + "(" + Trim(Str(trs!Length)) + ")"
    End Select
    If trs!Is_Nullable = "YES" Then
       ASql = ASql + " NULL, "
    Else
       ASql = ASql + " NOT NULL, "
    End If
    trs.MoveNext
Loop
ASql = Left(ASql, Len(ASql) - 2) + ") ON [PRIMARY]"
Print #1, ASql

On Error GoTo errorfound

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
mconn.Open strConnect
mcmd1.ActiveConnection = mconn
Sql = "sp_helpindex " + tbname$
mcmd1.CommandText = Sql
Set mrs = mcmd1.Execute
If mrs.BOF Or mrs.EOF Then
Else
   Do While Not mrs.EOF
      If InStr(mrs!index_description, "primary key") Then
         astr$ = "alter table " + tbname$ + " add constraint " + mrs!index_name + " primary key "
         If InStr(mrs!index_description, " unique ") Then astr$ = astr$ + "unique "
         If InStr(mrs!index_description, " clustered ") Then astr$ = astr$ + "clustered "
         If InStr(mrs!index_description, " nonclustered ") Then astr$ = astr$ + "nonclustered "
         astr$ = astr$ + "(" + mrs!index_keys + ")"
         Print #1, astr$
     Else
        CrStr = "Create "
        If InStr(mrs!index_description, " unique ") Then CrStr = CrStr + "unique "
        If InStr(mrs!index_description, " clustered ") Then CrStr = CrStr + "clustered "
        If InStr(mrs!index_description, " nonclustered ") Then CrStr = CrStr + "nonclustered "
        CrStr = CrStr + "index " + mrs!index_name + " on " + tbname$ + "(" + mrs!index_keys + ")"
        Print #1, CrStr
      End If
      mrs.MoveNext
   Loop
End If
mconn.Close
SkipForError:
On Error GoTo 0

If Option2.Value = 0 Then GoTo Done
LogComment = LogComment + " with data"
Sql = "Select * from " + tbname$
tcmd1.CommandText = Sql
Set trs = tcmd1.Execute
If trs.BOF Or trs.EOF Then GoTo Done
StartSql = "Insert into " + tbname$ + " ("
For z = 0 To trs.Fields.Count - 1
   n$ = trs.Fields(z).Name
   If InStr(n$, " ") Then n$ = "[" + n$ + "]"
   If trs.Fields(z).Type <> 128 Then StartSql = StartSql + n$ + ","
Next
StartSql = Left(StartSql, Len(StartSql) - 1) + ") values ("
Do While Not trs.EOF
   ISql = StartSql
   On Error GoTo errorfound2
   For z = 0 To trs.Fields.Count - 1
      Select Case trs.Fields(z).Type
         Case 129, 130, 202, 200, 135, 205, 203, 201, 72
            If IsNull(trs.Fields(z).Value) Then
               d$ = "''"
            Else
               If trs.Fields(z).Type = 135 Then
                  d$ = "'" + Str(trs.Fields(z).Value) + "'"
               Else
                  d$ = trs.Fields(z).Value
                  d$ = Replace(d$, "'", "''")
                  d$ = "'" + d$ + "'"
               End If
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
            d$ = Trim(hxstr)
         Case 20, 131, 5, 3, 6, 2, 17, 11, 4
            If IsNull(trs.Fields(z).Value) Then
               d$ = "0"
            Else
               d$ = Trim(Str(trs.Fields(z).Value))
            End If
      End Select
      ISql = ISql + d$ + ","
   Next
   ISql = Left(ISql, Len(ISql) - 1) + ")"
   Print #1, ISql
   trs.MoveNext
Loop
tconn.Close

Done:
Close
On Error GoTo 0
On Error GoTo errorProc
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + cmbServer.Text + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Open "c:\temp.sql" For Input As #1
rloop:
   Line Input #1, d$
   cmd1.CommandText = d$
   cmd1.Execute
   If Not EOF(1) Then GoTo rloop
Close
conn.Close
Kill "c:\temp.sql"
On Error GoTo 0
Call UpdateLog
Exit Sub

errorfound:
   Resume SkipForError

errorfound2:
   d$ = "''"
   Resume Next

errorProc:
   MsgBox "There was an error in the statement:" + d$
   Resume Next
   
End Sub

Private Sub Command2_Click()

Unload frmCopyTableAs

End Sub

Private Sub Form_Activate()

If Combo1.ListCount = 0 Then
   MsgBox "There are no tables in database " + DBName + "."
   Unload frmCopyTableAs
End If

End Sub

Private Sub Form_Load()

Text1.Text = ""
Option1.Value = 1
Combo1.Clear
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Set rs = conn.OpenSchema(adSchemaTables)
If rs.BOF Or rs.EOF Then
Else
   Do While Not rs.EOF
       If UCase(Left(rs!table_name, 3)) <> "SYS" Then
           Combo1.AddItem rs!table_name
       End If
       rs.MoveNext
   Loop
   conn.Close
End If
Combo1.Text = Combo1.List(0)
cmbServer.Clear
For z = 1 To xmax
   cmbServer.AddItem Xtras(z)
Next
cmbServer.Text = Server

End Sub
