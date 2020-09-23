VERSION 5.00
Begin VB.Form frmSearchShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4830
   ClientLeft      =   9645
   ClientTop       =   3540
   ClientWidth     =   5655
   Icon            =   "frmSearchShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5655
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSearchShow.frx":030A
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmSearchShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Text1.ForeColor = TForeColor
Text1.BackColor = TBackColor
Text1.FontName = TFontName
Text1.FontSize = TFontSize
Text1.Text = TextData

End Sub

