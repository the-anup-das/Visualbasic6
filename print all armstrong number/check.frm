VERSION 5.00
Begin VB.Form check 
   Caption         =   "Print all armstrong numbers "
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "Show Result"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox result 
      Height          =   1815
      Left            =   2040
      TabIndex        =   5
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox ending 
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox start 
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Ending Value"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Starting value :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim n As Integer
    st = Val(start.Text)
    ed = Val(ending.Text)
    
    For i = st To ed Step 1
        a = i
        s = 0
    
       While a > 0
            r = a Mod 10
            s = s + r ^ 3
            a = a \ 10
          
        Wend
        
           If s = i Then
                 result.Text = result.Text + Str(i)
            End If
            
    Next
    
End Sub
