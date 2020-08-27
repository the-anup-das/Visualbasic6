VERSION 5.00
Begin VB.Form Loancal 
   Caption         =   "Loan Calculater"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox text4 
      Height          =   405
      Left            =   3000
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton ShowPayment 
      Caption         =   "Show Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CheckBox Payearly 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Check if early payments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   3
      Top             =   2220
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Duration(in months)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Interest rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   900
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Loan amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Loancal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Payearly_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then ShowPayment.SetFocus

End Sub

Private Sub Text1_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_keyup(keycode As Integer, shift As Integer)
    If keycode = 13 Then Payearly.SetFocus
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub ShowPayment_Click()
    Dim payment As Single
    Dim loanirate As Single
    Dim loanduration As Integer
    Dim loanamount As Integer
    
    If IsNumeric(Text1.Text) Then
        loanamount = Text1.Text
    Else
        MsgBox "Please enter a valid amount"
        Exit Sub
    End If
    If IsNumeric(Text2.Text) Then
        loanirate = 0.01 * Text2.Text / 12
    Else
        MsgBox "Invalid interest rate, Please re-enter "
        Exit Sub
    End If
    If IsNumeric(Text3.Text) Then
        loanduration = Text3.Text
    Else
        MsgBox "Please specify the loan's duratin_as a number of month"
        Exit Sub
    End If
    
    payment = Pmt(loanirate, loanduration, -loanamount, 0, Payearly.Value)
    text4.Text = Format$(payment, "#.00")
    
End Sub
