VERSION 5.00
Begin VB.Form abbrition 
   Caption         =   "Abbribetion Of a String"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
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
      Height          =   735
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox result 
      Height          =   1095
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox nametxt 
      Height          =   1095
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
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
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Name :"
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
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "abbrition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim st As String, ln As Integer, k(100) As Integer, l As Integer
    st = nametxt.Text
    ln = Len(st)
    j = 1
    
    For i = ln To 1 Step -1
        If Mid(st, i, 1) = " " Then
            k(j) = i
            j = j + 1
        End If
    Next i
    
    For i = j To 2 Step -1
         result.Text = result.Text & Mid(st, k(i) + 1, 1) + "."
        Next i

   
        result.Text = result.Text + Mid(st, k(1), ln + 1 - k(1))
End Sub
