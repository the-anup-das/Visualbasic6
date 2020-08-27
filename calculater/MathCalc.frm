VERSION 5.00
Begin VB.Form MathCalc 
   Caption         =   "Math Calculater"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Display 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton equals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton div 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   17
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton over 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton times 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton plusminus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton dot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton c 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "MathCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Operand1 As Double, operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean
Private Sub c_Click()
    Display.Caption = ""
End Sub

Private Sub Digits_Click(Index As Integer)
    If ClearDisplay Then
        Disply.Caption = ""
        ClearDisplay = False
    End If
    Display.Caption = Display.Caption + Digits(Index).Caption
End Sub

Private Sub div_Click()
Operand1 = Val(Display.Caption)
    Operator = "/"
    Display.Caption = ""
End Sub

Private Sub dot_Click()
    If InStr(Display.Caption, ".") Then
        Exit Sub
    Else
        Display.Caption = Display.Caption + "."
    End If
End Sub

Private Sub equals_Click()
    Dim result As Double
    
    On Error GoTo ErrorHandler
    operand2 = Val(Display.Caption)
    If Operator = "+" Then result = Operand1 + operand2
    If Operator = "-" Then result = Operand1 - operand2
    If Operator = "*" Then result = Operand1 * operand2
    If Operator = "/" And operand2 <> "0" Then _
        result = Operand1 / operand2
    Display.Caption = result
    ClearDisplay = True
    Exit Sub
    
ErrorHandler:
    MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description
    Display.Caption = "ERROR"
    ClearDisplay = True
End Sub

Private Sub minus_Click()
Operand1 = Val(Display.Caption)
    Operator = "-"
    Display.Caption = ""
End Sub

Private Sub over_Click()
If Val(Display.Caption) <> 0 Then Display.Caption = 1 / Val(Display.Caption)
End Sub

Private Sub plus_Click()
    Operand1 = Val(Display.Caption)
    Operator = "+"
    Display.Caption = ""
End Sub

Private Sub plusminus_Click()
    Display.Caption = -Val(Display.Caption)
End Sub

Private Sub times_Click()
Operand1 = Val(Display.Caption)
    Operator = "*"
    Display.Caption = ""
    
End Sub
