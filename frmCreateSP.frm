VERSION 5.00
Begin VB.Form frmCreateSP 
   Caption         =   "Create Stored Procedure"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDB 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text"
      Height          =   3495
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   7455
      Begin VB.TextBox txtSP 
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Text            =   "frmCreateSP.frx":0000
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameter"
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   7455
      Begin VB.TextBox txtSize 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3000
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Size"
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Type"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Variable"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtSPName 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Database"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Stored Proc Name"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmCreateSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vars(500, 3) As String
Dim vmax As Integer
Dim proctext As String

Private Sub Command1_Click()

vmax = vmax + 1
Vars(vmax, 1) = txtVar.Text
Vars(vmax, 2) = Combo1.Text
Vars(vmax, 3) = txtSize.Text
Call UpdateSPText

End Sub

Private Sub Command2_Click()

On Error GoTo errorfound

errorhappened = 0
Set conn = New ADODB.Connection
Set cmd1 = New ADODB.Command
strConnect = "Provider=SQLOLEDB;server=" + Server + ";database=" + cmbDB.Text + ";uid=" + UID + ";pwd=" + PWD + ";"
conn.Open strConnect
cmd1.ActiveConnection = conn
cmd1.CommandText = txtSP.Text
cmd1.Execute
On Error GoTo 0
If errorhappened = 0 Then MsgBox "Stored Procedure Created."
Exit Sub

errorfound:
   errorhappened = 1
   MsgBox "Error - " + Err.Description
   Resume Next

End Sub

Private Sub Command3_Click()

Unload frmCreateSP

End Sub

Private Sub Form_Load()

txtSP.Text = ""
Combo1.Clear
Combo1.AddItem "bigint"
Combo1.AddItem "binary"
Combo1.AddItem "bit"
Combo1.AddItem "char"
Combo1.AddItem "datetime"
Combo1.AddItem "decimal"
Combo1.AddItem "float"
Combo1.AddItem "image"
Combo1.AddItem "int"
Combo1.AddItem "money"
Combo1.AddItem "nchar"
Combo1.AddItem "ntext"
Combo1.AddItem "numeric"
Combo1.AddItem "nvarchar"
Combo1.AddItem "real"
Combo1.AddItem "smalldatetime"
Combo1.AddItem "smallint"
Combo1.AddItem "smallmoney"
Combo1.AddItem "text"
Combo1.AddItem "timestamp"
Combo1.AddItem "tinyint"
Combo1.AddItem "uniqueidentifier"
Combo1.AddItem "varbinary"
Combo1.AddItem "varchar"
Combo1.Text = Combo1.List(0)

cmbDB.Clear
For z = 0 To frmMain!List1.ListCount - 1
   cmbDB.AddItem frmMain!List1.List(z)
   If frmMain!List1.Selected(z) Then cmbDB.Text = frmMain!List1.List(z)
Next

End Sub

Private Sub txtSP_Change()

pos = InStr(txtSP.Text, "AS")
If pos = 0 Then Exit Sub
proctext = Mid(txtSP.Text, pos + 2)

End Sub

Private Sub txtSPName_Change()

Call UpdateSPText

End Sub

Public Sub UpdateSPText()

sp$ = "/* Created by " + LogUser + " on " + Date$ + " */" + vbCrLf
sp$ = sp$ + "CREATE PROCEDURE dbo." + txtSPName.Text + vbCrLf
For z = 1 To vmax
   sp$ = sp$ + "@" + Vars(z, 1) + " " + Vars(z, 2)
   If Vars(z, 3) <> "" Then sp$ = sp$ + " (" + Vars(z, 3) + ")"
   If z <> vmax Then sp$ = sp$ + ","
   sp$ = sp$ + vbCrLf
Next
sp$ = sp$ + "AS" + proctext
txtSP.Text = sp$

End Sub

Private Sub txtVar_Change()

txtSize.Text = ""

End Sub
