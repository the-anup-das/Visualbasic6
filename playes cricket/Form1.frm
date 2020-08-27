VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      Height          =   1215
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "WESTINDIES"
      Height          =   735
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SRILANKA"
      Height          =   735
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ENGLAND"
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "AUSTRALIA"
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "PAKISTAN"
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "INDIA"
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Command1_Click()
MsgBox (Option1(n).Caption & " country plays test cricket")
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(Index) = True Then
n = Index
End If
End Sub
