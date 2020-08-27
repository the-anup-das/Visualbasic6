VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER TEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String, ch As String, str1 As String
Dim i As Integer, l As Integer, j As Integer
Private Sub Command1_Click()
    str = Text1.Text
    str1 = ""
    'Text1.Text = ""
    l = Len(str)
    For i = 1 To l Step 1
        ch = Mid(str, i, 1)
        
        If Asc(ch) >= 97 And Asc(ch) <= 122 Then
            j = Asc(ch) - 32
            str1 = str1 & Chr(j)
        Else
            str1 = str1 & ch
        End If
    Next i
    Text1.Text = str1
End Sub
