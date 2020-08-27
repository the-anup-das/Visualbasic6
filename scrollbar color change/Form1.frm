VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   615
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change(Index As Integer)
    Form1.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
End Sub
