VERSION 5.00
Begin VB.Form stopwatch 
   Caption         =   "Stop Watch"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   960
   End
   Begin VB.CommandButton stop 
      Caption         =   "Stop"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton start 
      Caption         =   "Start"
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
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label result 
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer, m As Integer
Private Sub start_Click()
    Timer1.Enabled = True
    s = 0
    m = 0
End Sub

Private Sub stop_Click()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If Timer1.Enabled = True Then
        s = s + 1
        result.Caption = str(m) & "m " & str(s) & "s"
        If s >= 60 Then
            m = m + 1
            s = 0
        End If
     End If
End Sub
