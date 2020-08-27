VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form3"
   ScaleHeight     =   7335
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7680
      Top             =   5760
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   1335
      Left            =   4200
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H008080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1335
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   1455
      Left            =   3600
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, y1 As Integer
Dim a As Boolean
Private Sub Form_Click()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    x = 5760
    y = 4220
    a = True
End Sub

Private Sub Timer1_Timer()
    x = x - 50
    Shape1.Top = x
    
    If x <= 1550 Then
        x = 5760
    End If
    
    If a = True Then
    y = y - 20
    Shape1.Left = y
    If y <= 3000 Then
        a = False
       ' y = 4320
        y1 = 3000
    End If
    Else
        y1 = y1 + 30
        Shape1.Left = y1
        If y1 >= 4220 Then
           y = 4220
            a = True
        End If
    End If
    
End Sub
