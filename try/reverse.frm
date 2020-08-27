VERSION 5.00
Begin VB.Form reverse 
   Caption         =   "Reverse a String"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "Click here to Reverse"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3960
      Width           =   5415
   End
   Begin VB.TextBox result 
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox strtxt 
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4095
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
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "reverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st, st1 As String
Private Sub show_Click()
    Dim st As String, ln, flag As Integer
    st = strtxt.Text
    ln = Len(st)
    For i = ln To 1 Step -1
         
       If Mid(st, i, 1) = " " Then
        flag = i
         For j = flag To ln Step 1
            st1 = st1 + Mid(st, j, 1)
        Next j
        ln = flag
      End If
  Next i
  
  For i = 1 To flag Step 1
    If Mid(st, i, 1) <> " " Then
       st1 = st1 + Mid(st, i, 1)
    End If
    Next i
    result.Text = st1
End Sub

