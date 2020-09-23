VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScript 
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Browse.."
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include Data"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Note:  Some scripting with include data might require a lot of time..."
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
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label lbStat 
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload frmScript

End Sub

Private Sub Command2_Click()

If Text1.Text = "" Then
   MsgBox "Missing Script File."
   Exit Sub
End If

frmScript.MousePointer = 11
Select Case ScriptType
   Case "DB"
      Open Text1.Text For Output As #1
      Call DoDB
   Case "Table"
      Open Text1.Text For Output As #1
      Call DoTable(TableName)
   Case "Proc"
      Open Text1.Text For Input As #1
      Call DoProc
End Select
frmScript.MousePointer = 1
Close
Unload frmScript

End Sub

Private Sub Command3_Click()

cd$ = CurDir
If ScriptType = "Proc" Then
    CommonDialog1.ShowOpen
Else
    CommonDialog1.ShowSave
End If

If CommonDialog1.FileName <> "" Then Text1.Text = CommonDialog1.FileName
ChDir cd$

End Sub

Private Sub Form_Load()

Text1.Text = "c:\test.sql"
If ScriptType = "Proc" Then
   Command2.Caption = "Process"
   Check1.Visible = False
   Label1.Caption = "Saved Script File"
Else
   Command2.Caption = "Generate"
   Check1.Visible = True
   Label1.Caption = "Script File to Save"
End If


End Sub

Public Sub DoDB()

Print #1, "Create Database " + DBName
Print #1, "Use " + DBName
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
Set rs = conn.OpenSchema(adSchemaTables)
Do While Not rs.EOF
    If UCase(Left(rs!table_name, 3)) <> "SYS" Then
        Call DoTable(rs!table_name)
    End If
    rs.MoveNext
Loop
conn.Close

End Sub
Public Sub DoTable(tbname$)

Set tconn = New ADODB.Connection
Set tcmd1 = New ADODB.Command
Set mconn = New ADODB.Connection
Set mcmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
tconn.Open strConnect
tcmd1.ActiveConnection = tconn
lbStat.Caption = "Processing Table " + tbname$
lbStat.Refresh

ASql = "Create Table dbo." + tbname$ + " ("
Sql = "sp_columns [" + tbname$ + "]"
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

If Check1.Value = 0 Then Exit Sub

Sql = "Select * from " + tbname$
tcmd1.CommandText = Sql
Set trs = tcmd1.Execute
If trs.BOF Or trs.EOF Then Exit Sub
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
On Error GoTo 0
Exit Sub

errorfound:
   Resume SkipForError

errorfound2:
   d$ = "''"
   Resume Next
   
End Sub
Public Sub DoProc()

Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
On Error GoTo errorfound

rloop:
   Line Input #1, Sql
   lbStat.Caption = Sql
   lbStat.Refresh
   cmd1.CommandText = Sql
   cmd1.Execute
   If Not EOF(1) Then GoTo rloop
   Close
conn.Close
Exit Sub

errorfound:
   MessageData = "There was a problem with the following statement:" + vbCrLf + Sql + vbCrLf + "Error-" + Err.Description
   Resume Next
   
End Sub

