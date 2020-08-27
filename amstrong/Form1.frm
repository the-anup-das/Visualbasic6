VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Form1.frx":0004
      Left            =   360
      List            =   "Form1.frx":0006
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Non-Amstrong"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Amstrong"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Number :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, r, s, b, a As Integer
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
n = Val(Text1.Text)
a = n
s = 0
While a > 0
    r = a Mod 10
    s = s + r * r * r
    a = a \ 10
Wend
If n = s Then
    
    cn.Execute ("insert into armstrong values('" & n & "')")
    List1.AddItem n
Else
    cn.Execute ("insert into nonarmstrong values('" & n & "')")
    List2.AddItem n
End If
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rn = New ADODB.Recordset
    cn.Open ("checkarm")
End Sub
