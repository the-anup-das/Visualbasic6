VERSION 5.00
Begin VB.Form tring 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      Index           =   2
      Left            =   3000
      Max             =   255
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      Index           =   1
      Left            =   1560
      Max             =   255
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   375
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   1320
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   4440
      Max             =   255
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   2520
      Max             =   255
      TabIndex        =   3
      Top             =   3360
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   600
      Max             =   255
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton show 
      Caption         =   "Click Here "
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   3720
      Width           =   5895
   End
   Begin VB.TextBox result 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Shape yellow 
      BorderColor     =   &H8000000E&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape green 
      BorderColor     =   &H8000000E&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   4500
      Shape           =   3  'Circle
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape red 
      BorderColor     =   &H8000000E&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   4440
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "tring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
result.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label1.Caption = Str(HScroll1.Value) + Str(HScroll2.Value) + Str(HScroll3.Value)
End Sub


Private Sub HScroll2_Change()
result.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label1.Caption = Str(HScroll1.Value) + Str(HScroll2.Value) + Str(HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
result.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label1.Caption = Str(HScroll1.Value) + Str(HScroll2.Value) + Str(HScroll3.Value)
End Sub


Private Sub HScroll4_Change(Index As Integer)
result.BackColor = RGB(HScroll4(0).Value, HScroll4(1).Value, HScroll4(2).Value)
End Sub

Private Sub show_Click()
 result.BackColor = RGB(205, 215, 55)
End Sub

Private Sub Timer1_Timer()
    If red.Visible = True Then
        red.Visible = False
        yellow.Visible = True
        green.Visible = False
    ElseIf yellow.Visible = True Then
        red.Visible = False
        yellow.Visible = False
        green.Visible = True
        
    ElseIf green.Visible = True Then
        red.Visible = True
        yellow.Visible = False
        green.Visible = False
    End If
End Sub
