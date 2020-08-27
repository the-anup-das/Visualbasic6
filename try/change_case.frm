VERSION 5.00
Begin VB.Form change_case 
   Caption         =   "Change Case"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton lower 
      Caption         =   "Upper Case"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton upper 
      Caption         =   "Lower Case"
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
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox result 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox strtxt 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter string"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "change_case"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lower_Click()
 Dim st, rev As String
     Dim ascii As Integer
     st = strtxt.Text
     rev = Trim(st)
     ln = Len(rev)
     
     For i = 1 To ln Step 1
        ascii = Asc(Mid(rev, i, 1))
        If ascii > 92 And ascii < 123 Then
           result.Text = result.Text + Chr(ascii - 32)
        Else
            result.Text = result.Text + Chr(ascii)
        End If
        
        Next i
End Sub

Private Sub upper_Click()
    Dim st, rev As String
     Dim ascii As Integer
     st = strtxt.Text
     rev = Trim(st)
     ln = Len(rev)
     
     For i = 1 To ln Step 1
        ascii = Asc(Mid(rev, i, 1))
        If ascii > 64 And ascii < 92 Then
           result.Text = result.Text + Chr(ascii + 32)
        Else
            result.Text = result.Text + Chr(ascii)
        End If
        
        Next i
End Sub
