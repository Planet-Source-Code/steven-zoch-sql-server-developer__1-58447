VERSION 5.00
Begin VB.Form frmViewCols 
   Caption         =   "View Cols Sorted"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4035
   Icon            =   "frmViewCols.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   7080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double click to select"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmViewCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload frmViewCols

End Sub

Private Sub Command2_Click()

Printer.Print frmMain.Caption
For z = 0 To List1.ListCount - 1
   Printer.Print List1.List(z)
Next
Printer.EndDoc

End Sub

Private Sub Form_Load()

List1.Clear
If TableOrCol = "Col" Then
   For z = 0 To frmMain!List3.ListCount - 1
      If Left(frmMain!List3.List(z), 1) = "=" Then
         Exit For
      End If
      List1.AddItem frmMain!List3.List(z)
   Next
Else
   For z = 0 To frmMain!List2.ListCount - 1
      If Left(frmMain!List2.List(z), 1) = "=" Then
         Exit For
      End If
      List1.AddItem frmMain!List2.List(z)
   Next
End If

End Sub

Private Sub List1_DblClick()

flag% = 0
For z = 0 To List1.ListCount - 1
   If List1.Selected(z) Then
      TableOrCol = List1.List(z)
      flag% = 1
      Exit For
   End If
Next
If flag% Then Unload frmViewCols

End Sub
