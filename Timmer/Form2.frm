VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   360
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3720
      Top             =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2400
      X2              =   2280
      Y1              =   2280
      Y2              =   1680
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
 Line1.X1 = Val(Line1.X1) + 20
 'Line1.Y1 = Val(Line1.Y1) - 20
 Label1.Caption = Val(Line1.X1)
 Label2.Caption = Val(Line1.X2)
 
 
 
End Sub
