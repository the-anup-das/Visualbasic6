VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "                         Calculator"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton equal 
      Caption         =   "="
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton div 
      Caption         =   "/"
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton over 
      Caption         =   "1/X"
      Height          =   495
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton times 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton plusminus 
      Caption         =   "+/-"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton clrbutton 
      Caption         =   "C"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton point 
      Caption         =   "."
      Height          =   495
      Index           =   10
      Left            =   1920
      TabIndex        =   0
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   1200
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   1920
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox display 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operand1 As Double, operand2 As Double, result As Double
Dim operator As String

Private Sub clrbutton_Click()
display.Text = ""
End Sub


Private Sub Digits_Click(Index As Integer)
    display.Text = display.Text + Digits(Index).Caption
End Sub

Private Sub minus_Click()
    operator = "-"
    operand1 = Val(display.Text)
    display.Text = ""
End Sub

Private Sub over_Click()
    If Val(display.Caption) <> 0 Then display.Caption = 1 / Val(display.Caption)
End Sub

Private Sub plus_Click()
    operator = "+"
    operand1 = Val(display.Text)
    display.Text = ""
End Sub

Private Sub plusminus_Click()
    If Val(display.Text) < 0 Then
        display.Text = -Val(display.Text)
    End If
End Sub

Private Sub times_Click()
    operator = "*"
    operand1 = Val(display.Text)
    display.Text = ""
End Sub
Private Sub div_Click()
     operator = "/"
    operand1 = Val(display.Text)
    display.Text = ""
End Sub
Private Sub point_Click(Index As Integer)
    If InStr(display.Text, ".") Then
        Exit Sub
    Else
        display.Text = display.Text + "."
    End If
End Sub
Private Sub equal_Click()
operand2 = Val(display.Text)
Select Case operator
    Case "+"
       
        result = operand1 + operand2
        display.Text = result 'Str()
    Case "-"
        
        result = operand1 - operand2
        display.Text = result
    Case "*"
        result = operand1 * operand2
        display.Text = result
    Case "/"
        If Val(operand2) <> "0" Then
        result = operand1 / operand2
        display.Text = result
        End If
End Select
End Sub

