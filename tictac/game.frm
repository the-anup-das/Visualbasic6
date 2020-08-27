VERSION 5.00
Begin VB.Form game 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "v"
   ClientHeight    =   4830
   ClientLeft      =   5850
   ClientTop       =   2685
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "game.frx":0000
   ScaleHeight     =   4830
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Picture         =   "game.frx":36D2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Buxton Sketch"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1 As Integer, p2 As Integer
Public enter As Boolean, j As Integer, r As Integer
Sub check()
    If (Command1.Caption = "X" And Command2.Caption = "X" And Command3.Caption = "X") Or (Command4.Caption = "X" And Command5.Caption = "X" And Command6.Caption = "X") Or (Command7.Caption = "X" And Command8.Caption = "X" And Command9.Caption = "X") Or (Command1.Caption = "X" And Command5.Caption = "X" And Command9.Caption = "X") Or (Command3.Caption = "X" And Command5.Caption = "X" And Command7.Caption = "X") Or (Command1.Caption = "X" And Command4.Caption = "X" And Command7.Caption = "X") Or (Command2.Caption = "X" And Command5.Caption = "X" And Command8.Caption = "X") Or (Command3.Caption = "X" And Command6.Caption = "X" And Command9.Caption = "X") Then
        'MsgBox ("p1 winner")
        win.Show
        game.Visible = False
        j = 3
        win.Label1.Caption = namefrm.p1name.Text + " Win "
        ElseIf (Command1.Caption = "o" And Command2.Caption = "o" And Command3.Caption = "o") Or (Command4.Caption = "o" And Command5.Caption = "o" And Command6.Caption = "o") Or (Command7.Caption = "o" And Command8.Caption = "o" And Command9.Caption = "o") Or (Command1.Caption = "o" And Command5.Caption = "o" And Command9.Caption = "o") Or (Command3.Caption = "o" And Command5.Caption = "o" And Command7.Caption = "o") Or (Command1.Caption = "o" And Command4.Caption = "o" And Command7.Caption = "o") Or (Command2.Caption = "o" And Command5.Caption = "o" And Command8.Caption = "o") Or (Command3.Caption = "o" And Command6.Caption = "o" And Command9.Caption = "o") Then
            'MsgBox ("p2 winner")
            win.Show
            game.Visible = False
            j = 3
                win.Label1.Caption = namefrm.p2name.Text + " Win "
            ElseIf (j = 9) Then
               i = MsgBox("Game Draw ", vbOKOnly, "DraW!!!")
               confirm.Show
               game.Visible = False
            End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "" Then
    If enter = False Then
     Command1.Caption = "X"
        Command1.BackColor = RGB(0, 0, 255)
        enter = True
    Else
    Command1.Caption = "o"
        Command1.BackColor = RGB(255, 0, 0)
     enter = False
     End If
     j = j + 1
     check
End If
End Sub



Private Sub Command10_Click()
game.Visible = False
pause.Show
End Sub

Private Sub Command2_Click()
If Command2.Caption = "" Then
    If enter = False Then
    Command2.Caption = "X"
    Command2.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command2.Caption = "o"
         Command2.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "" Then
     If enter = False Then
    Command3.Caption = "X"
    Command3.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command3.Caption = "o"
         Command3.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Command4_Click()
If Command4.Caption = "" Then
      If enter = False Then
    Command4.Caption = "X"
    Command4.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command4.Caption = "o"
         Command4.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Command5_Click()
If Command5.Caption = "" Then
     If enter = False Then
    Command5.Caption = "X"
    Command5.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command5.Caption = "o"
         Command5.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Command6_Click()
If Command6.Caption = "" Then
     If enter = False Then
    Command6.Caption = "X"
    Command6.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command6.Caption = "o"
         Command6.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
        End If
        check
    End If
End Sub

Private Sub Command7_Click()
If Command7.Caption = "" Then
     If enter = False Then
    Command7.Caption = "X"
    Command7.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command7.Caption = "o"
         Command7.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Command8_Click()
If Command8.Caption = "" Then
     If enter = False Then
    Command8.Caption = "X"
    Command8.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command8.Caption = "o"
         Command8.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
        End If
        check
End If
End Sub

Private Sub Command9_Click()
If Command9.Caption = "" Then
     If enter = False Then
    Command9.Caption = "X"
    Command9.BackColor = RGB(0, 0, 255)
     j = j + 1
     enter = True
     Else
        Command9.Caption = "o"
         Command9.BackColor = RGB(255, 0, 0)
          j = j + 1
        enter = False
    End If
    check
End If
End Sub

Private Sub Option1_Click()
enter = True
End Sub

Private Sub Timer1_Timer()
If Command1.Visible = True Then
  
    Command1.Visible = False
    Command4.Visible = Not True
    Command7.Visible = Not True
    Command4.Visible = Not True
    Command3.Visible = Not True
    Command5.Visible = Not True
    Command6.Visible = Not True
    Command9.Visible = Not True
Else
     Command1.Visible = True
    Command4.Visible = True
    Command7.Visible = True
    Command4.Visible = True
    Command3.Visible = True
    Command5.Visible = True
    Command6.Visible = True
    Command9.Visible = True
End If
End Sub
