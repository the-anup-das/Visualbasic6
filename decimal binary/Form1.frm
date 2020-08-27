VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset



Private Sub Form_Load()
 Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.jet.OLEDB.3.51"
     cn.Open App.Path & "\DECIMAL.mdb"
End Sub
Private Sub Command1_Click()
 Dim i, b, j, k As Integer
Dim arr(100) As Integer
Text2.Text = ""
 For i = 1 To 255 Step 1
    b = i
    j = 0
    res = 0
    For k = 0 To 9 Step 1
        arr(k) = 0
    Next
    While (b <> 0)
     r = b Mod 2
     arr(j) = r
     j = j + 1
     b = b \ 2
    Wend
 
    Text2.Text = ""
    k = 9
    While (k >= 0)
     Text2.Text = Text2.Text + Str(arr(k))
     k = k - 1
    Wend
    Text1.Text = Str(i)
    
      cn.Execute ("insert into Convert Values('" & Text1.Text & "','" & Text2.Text & "')")
    'cn.Execute ("insert into Login Values('" & usrText.Text & "','" & passText.Text & "')")
   
 Next
End Sub
