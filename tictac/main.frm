VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4920
   ClientLeft      =   3555
   ClientTop       =   3555
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0000
   ScaleHeight     =   4920
   ScaleWidth      =   8700
   Begin VB.Menu mnuoption 
      Caption         =   "&Option"
      Begin VB.Menu submnunew 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu submnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&Help"
      Begin VB.Menu submnuabout 
         Caption         =   "&About "
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub submnuabout_Click()
aboutme.Show
End Sub

Private Sub submnuexit_Click()
main.Visible = False
exitfrm.Show
End Sub

Private Sub submnunew_Click()
reset
game.Show
End Sub
