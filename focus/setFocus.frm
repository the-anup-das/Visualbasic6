VERSION 5.00
Begin VB.Form setFocus 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton savebttn 
      Caption         =   "Save Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1600
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "setFocus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub savebttn_Click()
    MsgBox "Record Saved. Click OK to enter another"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text1.setFocus
    
End Sub

Private Sub Text1_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text2.setFocus

End Sub

Private Sub Text2_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text3.setFocus
End Sub

Private Sub Text3_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text4.setFocus

End Sub

Private Sub Text4_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text5.setFocus

End Sub

Private Sub Text5_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then savebttn.setFocus

End Sub
