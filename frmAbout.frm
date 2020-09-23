VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About SQL Developer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Written by Steven Zoch. email eureka@"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.1.0"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SQL Developer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload frmAbout

End Sub

Private Sub Form_Load()

msg = "Written by Steven Zoch, email eureka@datasecuritysolutions.com." + vbCrLf
msg = msg + "This program is opensource and you are feel to modify, improve and add features.  You are not allowed to sell this product."

Label3.Caption = msg

End Sub
