VERSION 5.00
Begin VB.Form namefrm 
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   5070
   ClientTop       =   1920
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   Picture         =   "name.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton start 
      BackColor       =   &H8000000C&
      Height          =   975
      Left            =   2880
      Picture         =   "name.frx":4E7C2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox p2name 
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox p1name 
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player2 Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label p1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player1 Name "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "namefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub start_Click()
game.Show
namefrm.Visible = False
End Sub
