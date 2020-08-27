VERSION 5.00
Begin VB.Form find 
   Caption         =   "Finding a character "
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox result 
      Height          =   735
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton showbttn 
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
      Left            =   1560
      TabIndex        =   4
      Top             =   3720
      Width           =   4215
   End
   Begin VB.TextBox find 
      Height          =   735
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Height          =   735
      HideSelection   =   0   'False
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Find character "
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
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your text "
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
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub find_Change()
    If Len(find.Text) = 1 Then
       showbttn.SetFocus
    End If
End Sub


Private Sub showbttn_Click()

            Dim st As String, f As String, r As Integer, arr(100) As Integer
            
                 st = txt.Text
                 f = find.Text
                 s = Len(st)
                 c = 0
                 For i = 1 To s Step 1
                     If Mid(st, i, 1) = f Then
                         arr(c) = i
                         c = c + 1
                     End If
                     
                 Next i
                 
                 result.Text = "the character is found" + vbNewLine + Str(c) + " times and there positions are :"
               
                 
                For i = 0 To c - 1 Step 1
                     result.Text = result.Text + Str(arr(i)) + ","
                 Next i
End Sub
