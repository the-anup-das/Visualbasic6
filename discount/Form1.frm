VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Amount After Discount"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Discount"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Purchase amount"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sales As Double, discount As Double
Private Sub Command1_Click()
    sales = Val(Text1.Text)
    If sales <= 500 Then
        Text2.Text = "No Discount"
        Text3.Text = sales
    ElseIf sales > 500 And sales <= 1500 Then
        discount = sales * 0.05
        Text2.Text = discount
        Text3.Text = sales - discount
    ElseIf sales > 1500 And sales <= 3000 Then
        discount = sales * 0.15
        Text2.Text = discount
        Text3.Text = sales - discount
    ElseIf sales > 3000 And sales <= 5000 Then
        discount = sales * 0.2
        Text2.Text = discount
        Text3.Text = sales - discount
    Else
        discount = sales * 0.25
        Text2.Text = discount
        Text3.Text = sales - discount
    End If
End Sub
