VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Suger"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Rice"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Amount have to pay"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "RS-50/kg"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "RS-30/kg"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ITEM                         Countity                                     Price"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cost As Integer
Private Sub Command1_Click()
If Check1.Value = 1 Then
    cost = cost + (Val(Text1.Text)) * 30
End If

If Check2.Value = 1 Then
    cost = cost + (Val(Text2.Text)) * 50
End If

Text3.Text = cost
End Sub
