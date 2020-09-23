Attribute VB_Name = "Module1"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Global LogUser As String
Global LogAction As String
Global LogComment As String

Global Server As String
Global UID As String
Global PWD As String
Global TForeColor As String
Global TBackColor As String
Global TFontName As String
Global TFontSize As String
Global NullReq As String
Global List1TBackColor As String
Global List2TBackColor As String
Global List3TBackColor As String
Global List4TBackColor As String
Global List5TBackColor As String
Global List1TForeColor As String
Global List2TForeColor As String
Global List3TForeColor As String
Global List4TForeColor As String
Global List5TForeColor As String
Global ColData(500, 4) As String
Global cmax As Long
Global LastFunction As String

Global DBName As String
Global TableName As String
Global CountSql As String
Global StopNow As Boolean
Global LogServer As String
Global DBLog As String

Global conn As ADODB.Connection
Global cmd1 As ADODB.Command

Global ScriptType As String
Global SearchType As String
Global SearchStr As String
Global SearchInProgress As Boolean
Global ToDo(50000, 2)

Global CurrRow As Long
Global CurrCol As Long
Global EditOrInsert As String
Global ReloadNeeded As Boolean
Global Statements(5000) As String
Global CurrentStatement As String
Global CurrentDBName As String

Global Select01InUse As Boolean
Global Select02InUse As Boolean
Global Select03InUse As Boolean
Global Select04InUse As Boolean
Global Select05InUse As Boolean
Global Select06InUse As Boolean
Global Select07InUse As Boolean
Global Select08InUse As Boolean
Global Select09InUse As Boolean
Global Select10InUse As Boolean

Global TableOrCol As String
Global TextData As String
Global IFStoredProc(500, 2) As String
Global IFSPMax As Integer
Global SPName As String
Global TRName As String
Global Xtras(50) As String
Global xmax As Integer

Public Function CheckDependencies(SqlStr$, msg$) As Long

On Error GoTo errorfound

msg$ = ""
CheckDependencies = 0
If InStr(LCase(SqlStr$), " table ") = 0 Then Exit Function

Set dconn = New ADODB.Connection
Set dcmd1 = New ADODB.Command
Set drs = New ADODB.Recordset
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + DBName + ";uid=" + UID + ";pwd=" + PWD + ";"
dconn.Open strConnect
dcmd1.ActiveConnection = dconn
cs$ = ""
For z = Len(SqlStr$) To 1 Step -1
   If Mid(SqlStr$, z, 1) = " " Then
      Exit For
   Else
      cs$ = Mid(SqlStr$, z, 1) + cs$
   End If
Next
Sql = "sp_depends " + cs$
dcmd1.CommandText = Sql
Set drs = dcmd1.Execute
If drs.BOF Or drs.EOF Then
Else
   CheckDependencies = 1
   Do While Not drs.EOF
      msg$ = msg$ + drs!Name + " " + drs!Type + vbCrLf
      drs.MoveNext
   Loop
End If
dconn.Close
On Error GoTo 0
Exit Function

errorfound:
   Resume Next

End Function

Public Sub UpdateLog()

   Set conn = New ADODB.Connection
   Set cmd1 = New ADODB.Command
   strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   Sql = "insert into SqlDeveloperLog (LogAction,LogUser,LogDate,LogComment) values ('" + LogAction + "', '" + LogUser + "', '" + Str(Now) + "', '" + LogComment + "')"
   cmd1.CommandText = Sql
   cmd1.Execute
   conn.Close

End Sub

Public Sub CheckForLogDeletions()

On Error GoTo errorfound

   Set conn = New ADODB.Connection
   Set cmd1 = New ADODB.Command
   strConnect = "Provider=SQLOLEDB;server=" + LogServer + ";database=" + DBLog + ";uid=" + UID + ";pwd=" + PWD + ";"
   conn.Open strConnect
   cmd1.ActiveConnection = conn
   Sql = "select * from SqlDeveloperLog"
   cmd1.CommandText = Sql
   Set rs = cmd1.Execute
   If rs.BOF Or rs.EOF Then
   Else
      Do While Not rs.EOF
         daydiff = DateDiff("d", rs!LogDate, Now)
         If daydiff > 45 Then
            Sql = "Delete from SqlDeveloperLog where LogAction='" + rs!LogAction + " ' and LogUser='" + rs!LogUser + "' and LogDate='" + Str(rs!LogDate) + "'"
            cmd1.CommandText = Sql
            cmd1.Execute
         End If
         rs.MoveNext
      Loop
   End If
   conn.Close
   On Error GoTo 0
   Exit Sub
   
errorfound:
    Resume Next
   
End Sub
