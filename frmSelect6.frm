VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSelect6 
   Caption         =   "Select"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   Icon            =   "frmSelect6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   8640
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8640
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   15266
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   11880
      TabIndex        =   0
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label lbStat 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9240
      Width           =   4695
   End
End
Attribute VB_Name = "frmSelect6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload frmSelect6

End Sub

Private Sub Command2_Click()

For z = 1 To flex.Cols - 1
   If flex.ColWidth(z) - 50 > 0 Then
      flex.ColWidth(z) = flex.ColWidth(z) - 50
   End If
Next

End Sub

Private Sub Command3_Click()

For z = 1 To flex.Cols - 1
   flex.ColWidth(z) = flex.ColWidth(z) + 50
Next

End Sub

Private Sub Form_Load()

frmSelect6.Caption = CurrentStatement
Select06InUse = True

flex.Redraw = False
flex.Visible = False
Dim sizes(500) As Integer
Dim smax As Integer
On Error GoTo errorfound

Set sconn = New ADODB.Connection
Set scmd1 = New ADODB.Command

strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + CurrentDBName + ";uid=" + UID + ";pwd=" + PWD + ";"
sconn.Open strConnect
scmd1.ActiveConnection = sconn
scmd1.CommandText = CurrentStatement
Set srs = scmd1.Execute
If srs.BOF Or srs.EOF Then
   frmSelect6.MousePointer = 1
   MsgBox "There were errors or no data returned."
   GoTo Done
End If
RecordSetReturned = True

If RecordSetReturned = False Then GoTo Done
If srs.Fields.Count = 0 Then
   MsgBox "There was no data returned."
   GoTo Done
End If

'do headers
flex.Cols = srs.Fields.Count + 1
flex.Row = 0
smax = 0
cr& = 0
For z = 0 To srs.Fields.Count - 1
   flex.Col = z + 1
   flex.CellFontName = TFontName
   flex.CellFontSize = TFontSize
   flex.Text = LCase(srs.Fields(z).Name)
   smax = smax + 1
   sizes(smax) = Len(flex.Text)
Next
'do data
Do While Not srs.EOF
   cr& = cr& + 1
   lbStat.Caption = "Reading" + Str(cr&)
   lbStat.Refresh
   flex.Rows = cr& + 1
   flex.Row = cr&
   For z = 0 To srs.Fields.Count - 1
      flex.Col = z + 1
      flex.CellFontName = TFontName
      flex.CellFontSize = TFontSize
      Select Case srs.Fields(z).Type
         Case 3, 5, 131, 2, 6, 17, 4, 20
            If IsNull(srs.Fields(z).Value) Then
               flex.Text = "0"
            Else
               flex.Text = Trim(Str(srs.Fields(z).Value))
            End If
         Case 204, 128
            acnv = StrConv(srs.Fields(z).Value, vbUnicode)
            If IsNull(acnv) Then
               hxstr = ""
            Else
               hxstr = "0x"
               For y = 1 To Len(acnv)
                  hxstr = hxstr + Right("00" + Hex$(Asc(Mid$(acnv, y, 1))), 2)
               Next
            End If
            flex.Text = Trim(hxstr)
         Case 129, 200, 135, 202, 203, 11, 72, 201, 205, 130
            If IsNull(srs.Fields(z).Value) Then
               flex.Text = ""
               If NullReq = "Y" Then
                  flex.Text = "Null"
               End If
            Else
               flex.Text = srs.Fields(z).Value
            End If
         Case Else
            MsgBox "Unable to resolve type " + Str(srs.Fields(z).Type) + " on column " + srs.Fields(z).Name
      End Select
      If Len(flex.Text) > sizes(z + 1) Then sizes(z + 1) = Len(flex.Text)
   Next
   srs.MoveNext
Loop

flex.ColWidth(0) = 0
For z = 1 To smax
   flex.ColWidth(z) = (sizes(z) + 1) * 100
Next
flex.Redraw = True
flex.Visible = True

Done:
frmSelect6.MousePointer = 1
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
sconn.Close

On Error GoTo 0
Exit Sub

errorfound:
   If Err.Number <> 6 Then MsgBox "There was the following error - " + Err.Description
   Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

Select06InUse = False

End Sub
