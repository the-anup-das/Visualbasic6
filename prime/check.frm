VERSION 5.00
Begin VB.Form check 
   Caption         =   "Check a number is prime or not"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton show 
      Caption         =   "Show Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub show_Click()
    Dim n, c, i As Integer
    n = Val(number.Text)
    c = 0
    
    For i = 1 To n Step 1
        If n Mod i = 0 Then
            c = c + 1
        End If
    Next i
    
    If c = 2 Then
        MsgBox "prime number"
    Else
        MsgBox "Non-prime number"
    End If
    
End Sub
