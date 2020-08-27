VERSION 5.00
Begin VB.Form check 
   Caption         =   "Check  a number is palindrome or not"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox result 
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox number 
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Show 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Result "
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Number"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Show_Click()
        Dim n As Integer, num As Integer, j As Integer, r As Integer
        n = Val(number.Text)
        num = n
        
        While num > 0
            r = num Mod 10
            j = j * 10 + r
            num = num \ 10
        Wend
            
            If n = j Then
                result.Text = "palindrome"
            Else
                result.Text = "non palindrome number"
            End If
End Sub

