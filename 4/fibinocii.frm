VERSION 5.00
Begin VB.Form fibinocii 
   Caption         =   "Fibinocii"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   FillColor       =   &H008080FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "fibinocii.frx":0000
   ScaleHeight     =   4920
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      BackColor       =   &H000000C0&
      Caption         =   "Click here"
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
      Left            =   240
      MaskColor       =   &H000080FF&
      Picture         =   "fibinocii.frx":98E1
      TabIndex        =   1
      Top             =   3840
      Width           =   5895
   End
   Begin VB.TextBox result 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "fibinocii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub show_Click()
    Dim a, b, f, c, n As Integer
    a = 0
    b = 1
    n = InputBox("how many fibinocii number you want??", "Input")
    result.Text = result.Text + Str(a) + " " + Str(b) + " "
    For i = 1 To n - 2 Step 1
        c = a + b
        a = b
        b = c
        result.Text = result.Text + Str(c) + " "
    Next i
End Sub
