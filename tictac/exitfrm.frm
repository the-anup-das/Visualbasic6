VERSION 5.00
Begin VB.Form exitfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   3045
   ClientLeft      =   5070
   ClientTop       =   4230
   ClientWidth     =   5955
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "exitfrm.frx":0000
   ScaleHeight     =   3045
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton no 
      Height          =   495
      Left            =   4440
      Picture         =   "exitfrm.frx":3D3F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton yes 
      Height          =   495
      Left            =   1080
      MaskColor       =   &H8000000B&
      Picture         =   "exitfrm.frx":3EC1C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "exitfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub yes_Click()
End
End Sub
Private Sub no_Click()
start.Show
exitfrm.Visible = False
End Sub



