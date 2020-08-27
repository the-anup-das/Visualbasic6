VERSION 5.00
Begin VB.Form pause 
   BorderStyle     =   0  'None
   ClientHeight    =   4920
   ClientLeft      =   3900
   ClientTop       =   3450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "pause.frx":0000
   ScaleHeight     =   4920
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton menu 
      Height          =   495
      Left            =   3960
      Picture         =   "pause.frx":7F2F
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton continue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Height          =   735
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "pause.frx":A471
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "pause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub continue_Click()
game.Show
pause.Visible = False
End Sub

Private Sub menu_Click()
main.Show
pause.Visible = False
End Sub
