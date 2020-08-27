VERSION 5.00
Begin VB.Form palindrome 
   Caption         =   "Check Palindrome string"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "Show Result"
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox result 
      Height          =   1095
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox strtxt 
      Height          =   1095
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
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
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "palindrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim st, st1 As String, ln, j, i As Integer, rev As String
    st1 = strtxt.Text
    st = Trim(st1)
    st = UCase(st1)
    ln = Len(st)
    j = 1
    
     For i = ln To 1 Step -1
        rev = rev + Mid(st, i, 1)
    Next i
                                          
        If StrComp(st, rev, 0) = 0 Then
            result.Text = "Palindrome string"
        Else
            result.Text = "Non-Palindrome string"
        End If
       
End Sub
