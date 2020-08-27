VERSION 5.00
Begin VB.Form question 
   Caption         =   "Draw"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   5775
   End
   Begin VB.TextBox result 
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "question"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim n, i, s, j As Integer
    n = InputBox("Enter how many lines you want?? ", "Input")
    s = InputBox("Enter starting point: ", "Input")
        
    For i = n To 1 Step -1
        For j = i To 1 Step -1
            result.Text = result.Text + Str(s)
        Next j
         s = s - 1
       result.Text = result.Text + vbNewLine
    Next i
End Sub
