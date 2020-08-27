VERSION 5.00
Begin VB.Form ques14 
   Caption         =   "Draw a triangle "
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7980
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
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   6855
   End
   Begin VB.TextBox result 
      Height          =   3015
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   6855
   End
End
Attribute VB_Name = "ques14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim n As Integer
    n = InputBox("How many lines you want?? ", "Input")
    
    For i = n To 1 Step -1
        For k = 0 To n - i Step 1
            result.Text = result.Text + " "
        Next
        
        For j = 2 * i - 1 To 1 Step -1
            result.Text = result.Text + "*"
        Next
        result.Text = result.Text + vbNewLine
    Next
End Sub
