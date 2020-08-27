VERSION 5.00
Begin VB.Form check 
   Caption         =   "Print all the even number"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox ending 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox starting 
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
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton show 
      Caption         =   "Show Result "
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
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
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Ending Value :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Value :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub show_Click()
    Dim i, st, ed As Integer
     st = Val(starting.Text)
     ed = Val(ending.Text)
     
     For i = st To ed Step 1
        If i Mod 2 = 0 Then
            result.Text = result.Text + Str(i) '+ vbNewLine
        End If
       Next
       
End Sub



Private Sub starting_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ending.SetFocus
End Sub
