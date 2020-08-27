VERSION 5.00
Begin VB.Form check 
   Caption         =   "Print all the prefect number between the range"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7185
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
      Height          =   735
      Left            =   1320
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox result 
      Height          =   1215
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox ending 
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox starting 
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Result "
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
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
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
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Starting Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
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

    Dim n, s, i, st, ed As Integer
    st = Val(starting.Text)
    ed = Val(ending.Text)
    
    For j = st To ed Step 1
        
        n = j
        s = 0
        i = 1
    
        While i < n
            If n Mod i = 0 Then
                s = s + i
            End If
                 i = i + 1
        Wend
        
        If n = s Then
            result.Text = result.Text + Str(n)
        End If
        
    Next j

End Sub
