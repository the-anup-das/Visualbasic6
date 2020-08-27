VERSION 5.00
Begin VB.Form Calculaters 
   Caption         =   "Calculater"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton play 
      Caption         =   "Play a Game"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton loan 
      Caption         =   "Loan Calculator"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton math 
      Caption         =   "Math Calculator"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Calculaters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Exit_Click()
    End
End Sub

Private Sub loan_Click()
    Loancal.Show
End Sub

Private Sub math_Click()
    MathCalc.Show
End Sub

Private Sub play_Click()
    play.Left = Rnd() * (Calculaters.Width - play.Width)
    play.Top = Rnd() * (Calculaters.Height - play.Height)
End Sub
