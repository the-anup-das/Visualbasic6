VERSION 5.00
Begin VB.Form amstng 
   Caption         =   "check number is either amstrong or not"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton result 
      Caption         =   "Show Result"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
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
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your number :"
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
      Width           =   2655
   End
End
Attribute VB_Name = "amstng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub result_Click()
    Dim n, r, a, s As Integer
    n = Val(number.Text)
    s = 0
    a = n
    
    While a > 0
        r = a Mod 10
        s = s + r ^ 3
        a = a \ 10
    Wend
    
        If s = n Then
            MsgBox "Number is Amstrong"
            
        Else
            MsgBox "Non Amstrong Number "
        End If
End Sub
