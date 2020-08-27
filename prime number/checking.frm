VERSION 5.00
Begin VB.Form checking 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Prime"
      Height          =   615
      Left            =   7320
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   6720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton show 
      Caption         =   "Show Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox number 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset



Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open ("typenu")
End Sub

Private Sub show_Click()
    Dim n, c, i As Integer
    n = Val(number.Text)
    c = 0
    
    For i = 1 To n Step 1
        If n Mod i = 0 Then
            c = c + 1
        End If
    Next i
    
    If c = 2 Then
        MsgBox "prime number"
        cn.Execute ("insert into prime values('" & number.Text & "')")
    Else
        MsgBox "Non-prime number"
        cn.Execute ("insert into nonprime values ('" & number.Text & "')")
    End If
    
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
    Set rs = cn.Execute("select * from prime")
    
    While Not rs.EOF
        Text1.Text = Text1.Text + Str(rs(0))
        rs.MoveNext
    Wend
End Sub
