VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   2640
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   4920
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H008080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1335
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, y1, z1, x As Integer
Dim b, a As Boolean
Private Sub Form_Load()
x1 = 0
y1 = 0
b = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
Shape1.Left = x
Shape1.Top = Y
End If
End Sub

Private Sub Timer1_Timer()
If a = True Then
x1 = x1 + 30
Shape1.Left = x1

If x1 >= Form1.ScaleWidth - 1400 Then
    a = False
 End If
 x = Form1.ScaleWidth - 1400
 Else
    x = x - 30
    Shape1.Left = x
    If x <= 0 Then
        x1 = 0
        a = True
    End If
End If

If b = True Then
 y1 = y1 + 30
 Shape1.Top = y1
 If y1 >= Form1.ScaleHeight - 1400 Then
    b = False
    End If
    z1 = Form1.ScaleHeight - 1400
Else

    z1 = z1 - 30
    Shape1.Top = z1
    If z1 <= 0 Then
        y1 = 0
        b = True
    End If
End If
End Sub

Private Sub Timer2_Timer()
    Shape1.FillColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
End Sub
