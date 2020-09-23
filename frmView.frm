VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmView 
   Caption         =   "Data"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   9720
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   9960
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flex2 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   17171
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload frmView

End Sub

Private Sub Command2_Click()

For z = 1 To flex2.Cols - 1
   If flex2.ColWidth(z) - 50 > 0 Then
      flex2.ColWidth(z) = flex2.ColWidth(z) - 50
   End If
Next

End Sub

Private Sub Command3_Click()

For z = 1 To flex2.Cols - 1
   flex2.ColWidth(z) = flex2.ColWidth(z) + 50
Next

End Sub

Private Sub Form_Load()

Dim cdat(5000, 2)

If SearchInProgress Then
On Error GoTo StopDraw
   flex2.Redraw = False
   flex2.Visible = False
   flex2.Cols = 99
   For z = 0 To frmSearch!List1.ListCount - 1
      t$ = frmSearch!List2.List(z)
      p = 0
      cm = 0
      For zz = 1 To Len(t$)
         If Mid(t$, zz, 1) = "|" Then
            p = p + 1
            If p = 3 Then p = 1
            If p = 1 Then cm = cm + 1
            cdat(cm, p) = ct$
            ct$ = ""
         Else
            ct$ = ct$ + Mid(t$, zz, 1)
         End If
      Next
      ml = 0
      cr& = cr& + 1
      flex2.Rows = cr& + 1
      flex2.Row = cr&
      For zz = 1 To cm
            flex2.Col = zz
            flex2.CellFontName = TFontName
            flex2.CellFontSize = TFontSize
            flex2.Text = cdat(zz, 1)
         Next
      cr& = cr& + 1
      flex2.Rows = cr& + 1
      flex2.Row = cr&
      For zz = 1 To cm
         flex2.Col = zz
         flex2.CellFontName = TFontName
         flex2.CellFontSize = TFontSize
         flex2.Text = cdat(zz, 2)
      Next
      cr& = cr& + 1
      flex2.Rows = cr&
   Next
   flex2.ColWidth(0) = 0
   flex2.Redraw = True
   flex2.Visible = True
End If
Okay:
On Error GoTo 0
Exit Sub

StopDraw:
   MsgBox "Data is too big to display."
   Resume Okay
   
End Sub
