VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   645
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   3840
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0004
      Left            =   840
      List            =   "Form1.frx":0006
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reverse String"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input string"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Inputed String"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "NON-PALINDROME"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "PALINDROME"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "REVERSE STRING"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String, str2 As String, ch As String
Dim l As Integer
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
    str1 = ""
    Text2.Text = ""
    str1 = InputBox("Enter your string")
    Text2.Text = str1
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
     Text1.Text = ""
    l = Len(str1)
    For i = l To 1 Step -1
       ch = Mid(str1, i, 1)
       str2 = str2 & ch
    Next i
    Text1.Text = str2
    
    If StrComp(str1, str2) = 0 Then
        MsgBox ("String is Palindrome")
        List1.AddItem str1
        cn.Execute ("insert into palindrome values('" & str1 & "')")
    Else
        MsgBox ("String is Non-palindrome")
        cn.Execute ("inset into nonpalindrome values('" & str1 & "')")
        List2.AddItem str1
    End If
        
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open ("palindromecheck")
End Sub
