VERSION 5.00
Begin VB.Form confirm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   6240
   ClientTop       =   3840
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   Picture         =   "confirm.frx":0000
   ScaleHeight     =   3585
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton conf 
      Height          =   615
      Left            =   720
      Picture         =   "confirm.frx":3644E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   480
      Picture         =   "confirm.frx":3C160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "confirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
reset
choose.Show
confirm.Visible = False
End Sub

Private Sub conf_Click()
confirm.Visible = False
exitfrm.Show
End Sub

Sub reset()
    j = 0
    game.Command1.Caption = ""
    game.Command2.Caption = ""
    game.Command3.Caption = ""
    game.Command4.Caption = ""
    game.Command5.Caption = ""
    game.Command6.Caption = ""
    game.Command7.Caption = ""
    game.Command8.Caption = ""
    game.Command9.Caption = ""
    game.Command1.BackColor = RGB(255, 255, 255)
    game.Command2.BackColor = RGB(255, 255, 255)
    game.Command3.BackColor = RGB(255, 255, 255)
    game.Command4.BackColor = RGB(255, 255, 255)
    game.Command5.BackColor = RGB(255, 255, 255)
    game.Command6.BackColor = RGB(255, 255, 255)
    game.Command7.BackColor = RGB(255, 255, 255)
    game.Command8.BackColor = RGB(255, 255, 255)
     game.Command9.BackColor = RGB(255, 255, 255)
End Sub


