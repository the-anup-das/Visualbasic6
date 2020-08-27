VERSION 5.00
Begin VB.Form check 
   Caption         =   "Check a number is Even or Odd"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton result 
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
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Number  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   2535
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub number_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then result.SetFocus
End Sub

Private Sub result_Click()
    Dim n, a As Integer
    n = Val(number.Text)
    a = n Mod 2
    
    If a = 0 Then
        MsgBox "Even Number"
        
    Else
        MsgBox "odd number "
   End If
End Sub
