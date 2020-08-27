VERSION 5.00
Begin VB.Form choose 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   5265
   ClientTop       =   1350
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "choose.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      Picture         =   "choose.frx":7B15A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      Picture         =   "choose.frx":7E19C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
   End
End
Attribute VB_Name = "choose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
reset
namefrm.Show
choose.Visible = False
End Sub

Public Sub Option1_Click()
If Option1.Value = True Then
    game.enter = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    game.enter = True
End If
End Sub

Private Sub start_Click()

End Sub
