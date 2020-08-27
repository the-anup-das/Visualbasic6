VERSION 5.00
Begin VB.Form astng 
   Caption         =   "Check a Number is Armstrong or  Not"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
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
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Result :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter  a Number :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "astng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim n, a, s, r As Integer
    n = Val(number.Text)
    a = n
    s = 0
    
    While a > 0
        r = a Mod 10
        s = s + r ^ 3
        a = a \ 10
        
    Wend
        
        If s = n Then
           result.Text = " Amstrng number"
            
        Else
            result.Text = "Nonamstrong number"
        End If
            
                
End Sub
