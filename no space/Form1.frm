VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0           'this give no space on text
        MsgBox ("No space allowed")
    End If
End Sub
