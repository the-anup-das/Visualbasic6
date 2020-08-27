VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "                             Display even or odd number"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "ENTER NUMBER "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "THE NUMBER IS EVEN OR ODD"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
Dim r As Integer
    r = Text1.Text Mod 2
    If r = 0 Then
    Label3.Caption = Text1.Text & " is Even no"
    Else
    Label3.Caption = Text1.Text & " is Odd no"
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub

