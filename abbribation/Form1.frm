VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "convert"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your Text"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String, l As Integer, arr(20) As Integer
Private Sub Command1_Click()
    str1 = Text1.Text
    l = Len(str1)
    j = 1
    For i = l To 1 Step -1
        If Mid(str1, i, 1) = " " Then
                 arr(j) = i
                 j = j + 1
        End If
    Next i
    
    For i = j To 2 Step -1
        Text2.Text = Text2.Text + Mid(str1, arr(i) + 1, 1) + "."
    Next i
        Text2.Text = Text2.Text + Mid(str1, arr(1), l + 1 - arr(1))
End Sub
