VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions2 
   Caption         =   "List Boxes Color"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   Icon            =   "frmOptions2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7200
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Set"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Text"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Backgr"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Text"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Backgr"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Text"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Backgr"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Text"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Backgr"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Text"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backgr"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox List5 
      Height          =   1230
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List4 
      Height          =   1230
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "You will need to click Set again at the main option screen."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   2400
      Width           =   4335
   End
End
Attribute VB_Name = "frmOptions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List1.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub Command10_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List5.ForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command11_Click()

List1TBackColor = List1.BackColor
List1TForeColor = List1.ForeColor
List2TBackColor = List2.BackColor
List2TForeColor = List2.ForeColor
List3TBackColor = List3.BackColor
List3TForeColor = List3.ForeColor
List4TBackColor = List4.BackColor
List4TForeColor = List4.ForeColor
List5TBackColor = List5.BackColor
List5TForeColor = List5.ForeColor
Unload frmOptions2

End Sub

Private Sub Command12_Click()

Unload frmOptions2

End Sub

Private Sub Command2_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List1.ForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command3_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List2.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub Command4_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List2.ForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command5_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List3.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub Command6_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List3.ForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command7_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List4.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub Command8_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List4.ForeColor = CommonDialog1.Color
End If

End Sub

Private Sub Command9_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color Then
   List5.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub Form_Load()

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
keystr = ""

List1.AddItem "This"
List1.AddItem "is"
List1.AddItem "the"
List1.AddItem "Database"

List2.AddItem "This"
List2.AddItem "is"
List2.AddItem "the"
List2.AddItem "Tables"

List3.AddItem "This"
List3.AddItem "is"
List3.AddItem "the"
List3.AddItem "Columns"

List4.AddItem "This"
List4.AddItem "is"
List4.AddItem "the"
List4.AddItem "Stored Procedures"

List5.AddItem "This"
List5.AddItem "is"
List5.AddItem "the"
List5.AddItem "Triggers"

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


End Sub
