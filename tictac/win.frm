VERSION 5.00
Begin VB.Form win 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   4485
   ClientTop       =   765
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   Picture         =   "win.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4440
      Top             =   5280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Win "
      BeginProperty Font 
         Name            =   "SketchFlow Print"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   4215
   End
End
Attribute VB_Name = "win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Click()
confirm.Show
win.Visible = False
End Sub


Private Sub Timer1_Timer()
    Label1.ForeColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
End Sub
