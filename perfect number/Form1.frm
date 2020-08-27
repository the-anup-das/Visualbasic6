VERSION 5.00
Begin VB.Form check 
   Caption         =   "Check a number is perfect or not"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6285
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your number "
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
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim n, s, i As Integer
    n = Val(number.Text)
    s = 0
    i = 1
    
    While i < n
        If n Mod i = 0 Then
            s = s + i
        End If
             i = i + 1
    Wend
    
    If n = s Then
        result.Text = "perfect"
    Else
        result.Text = "non perfect"
    End If
End Sub
