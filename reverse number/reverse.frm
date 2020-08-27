VERSION 5.00
Begin VB.Form reverse 
   Caption         =   "Reverse a Number"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton click 
      Caption         =   "Click to Reverse"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox reverse 
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox number 
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Reverse Number"
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
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "reverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub click_Click()
                Dim num As String
                num = number.Text
                reverse.Text = StrReverse(num)
End Sub
