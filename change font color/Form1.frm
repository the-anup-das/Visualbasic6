VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      Left            =   120
      Max             =   255
      TabIndex        =   5
      Top             =   3120
      Width           =   3855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   2760
      Width           =   3855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   3
      Top             =   2400
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":0000
      Left            =   4440
      List            =   "Form1.frx":0002
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "TEXT COLOR CHANGE "
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
j = 10
For i = 0 To 20 Step 1
    j = j + 2
    List1.AddItem j, i
Next i
End Sub

Private Sub HScroll1_Change(Index As Integer)
 Text1.ForeColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
End Sub

Private Sub List1_Click()
    Text1.FontSize = List1.List(List1.ListIndex)
End Sub
