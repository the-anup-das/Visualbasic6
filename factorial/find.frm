VERSION 5.00
Begin VB.Form find 
   Caption         =   "Find Factorial of any number"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Result 
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
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox factorial 
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox number 
      Height          =   975
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Factorial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Result_Click()
            Dim n As Integer
            f = 1
            n = Val(number.Text)
            
            While n > 0
                f = f * n
                n = n - 1
            Wend
            
            factorial.Text = Str(f)
End Sub
