VERSION 5.00
Begin VB.Form frmcheck 
   Caption         =   "Checking entered number is armstrong or not"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Delete previous stored items"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to check"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   60
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your number:"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
    Dim num As Integer, s As Integer, r As Integer
    num = Val(Text1.Text)
    
    While (num > 0)
        r = num Mod 10
        s = s + r ^ 3
        num = num \ 10
    Wend
    
    If s = Val(Text1.Text) Then
        Label2.ForeColor = RGB(10, 255, 20)
        Label2.Caption = "Your entered number is Armstrong"
        cn.Execute ("insert into armstrong values('" & s & "')")
    Else
        Label2.Caption = "Your entered number is Non-Armstrong"
        cn.Execute ("insert into nonarmstrong values('" & s & "')")
    End If
End Sub

Private Sub Command2_Click()
    Dim s As String
    s = InputBox("Enter Password", "Security")
    
    If s = "Delete" Then
        On Error GoTo err
        cn.Execute ("delete from armstrong")
        cn.Execute ("delete from nonarmstrong")
        MsgBox "data are deleted"
        Exit Sub
    End If
err:
    MsgBox "can't delete" & Chr(13) & err.Description
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.Open ("armstrongcheck")
End Sub
