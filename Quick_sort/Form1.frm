VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Result:"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "How many numbers:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(10), n, l As Integer

Private Sub Command1_Click()
Dim i As Integer
n = Val(Text1.Text)
For i = 1 To n
    num = InputBox("Enter number")
    arr(i) = num
Next i
l = 0
Call quick_sort(l, n - 1)




End Sub

Public Sub quick_sort(ByRef low As Integer, ByRef high As Integer)
Dim pivot, left, right, temp As Integer

If low > high Then
   Return  Command1_Click
End If
    pivot = arr(low)
    left = low + 1
    right = hight
    
While (left <= right)
    While ((arr(left) <= pivot) And (left <= high))
        left = left + 1
    Wend
    While ((arr(right) > pivot) And (right >= low))
        right = right - 1
    Wend
    If left < right Then
        temp = arr(left)
        arr(left) = arr(right)
        arr(right) = temp
    End If
Wend

If low < right Then
    temp = arr(right)
    arr(right) = arr(low)
    arr(low) = temp
End If

Call quick_sort(low, right - 1)
Call quick_sort(right + 1, high)
End Sub

Private Sub Command2_Click()
Dim i As Integer
For i = 1 To n
    Text2.Text = arr(i) & " "
Next i
End Sub
