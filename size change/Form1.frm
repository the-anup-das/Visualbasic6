VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "10", 0
List1.AddItem "12", 1
List1.AddItem "14", 2
List1.AddItem "20", 3
List1.AddItem "30", 4
End Sub

Private Sub List1_Click()
Text1.FontSize = List1.List(List1.ListIndex)
End Sub

