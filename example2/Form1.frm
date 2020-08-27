VERSION 5.00
Begin VB.Form Example2 
   Caption         =   "TextBOx Demo"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Show Message"
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Date"
      Height          =   615
      Left            =   2940
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Text"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox text1 
      Height          =   2655
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Example2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    text1.Text = ""
End Sub

Private Sub Command2_Click()
    text1.Text = Date
End Sub

Private Sub Command3_Click()
    text1.Text = "welcome to @wsome "
End Sub

