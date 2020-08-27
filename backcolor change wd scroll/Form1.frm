VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      LargeChange     =   10
      Left            =   4440
      Max             =   255
      TabIndex        =   2
      Top             =   4200
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      LargeChange     =   10
      Left            =   2280
      Max             =   255
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      LargeChange     =   10
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
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
