VERSION 5.00
Begin VB.Form start 
   BorderStyle     =   0  'None
   ClientHeight    =   6975
   ClientLeft      =   4680
   ClientTop       =   1530
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   Picture         =   "start.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1320
      Top             =   6000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Any were to Continue "
      BeginProperty Font 
         Name            =   "Segoe Marker"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Click()
choose.Show
start.Visible = False
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = True Then
    Label1.Visible = False
Else
    Label1.Visible = True
End If
End Sub
