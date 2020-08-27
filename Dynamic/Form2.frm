VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   ScaleHeight     =   5790
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1920
      Top             =   4560
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H008080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1335
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, y1, z1, x As Integer
Dim b As Boolean, a As Boolean
Private Sub Form_Load()
x1 = 0
y1 = 0
b = True
a = True
End Sub


Private Sub Timer1_Timer()
If a = True Then
x1 = x1 + 30
Shape1.Top = x1

If x1 >= Form1.ScaleHeight - 1400 Then
    a = False
 End If
 x = Form1.ScaleHeight - 1400
 Else
    x = x - 30
    Shape1.Top = x
    
    If x <= 0 Then
        x1 = 0
    a = True
    End If
End If

If b = True Then
 y1 = y1 + 30
 Shape1.Left = y1
 If y1 >= Form1.ScaleWidth - 1400 Then
    b = False
    End If
    z1 = Form1.ScaleWidth - 1400
Else

    z1 = z1 - 30
    Shape1.Left = z1
    If z1 <= 0 Then
        y1 = 0
        b = True
    End If
End If
End Sub

Private Sub Timer2_Timer()
    Shape1.FillColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
End Sub

