VERSION 5.00
Begin VB.Form check 
   Caption         =   "Check a number is fibonacii or not"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "Show Result"
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
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox number 
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   3615
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
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
            Dim n, arr(1000), j, a, b, c As Integer
            n = Val(number.Text)
            j = 0
            c = 0
            a = 0
            b = 1
            arr(j) = a
            
            For i = 0 To n + 1 Step 1
            
                c = a + b
                a = b
                b = c
                j = j + 1
                arr(j) = c
                'result.Text = result.Text + Str(arr(j))
               
            Next
            
             j = 0
             While arr(j) <= n
                If n = arr(j) Then
                    result.Text = "Fibonacii number"
                    'result.Text = result.Text + Str(arr(j))
                Else
                    result.Text = "The number is not fibonacii number "
                   'result.Text = result.Text + Str(arr(j))
                End If
                j = j + 1
            Wend
  
            
                
End Sub
