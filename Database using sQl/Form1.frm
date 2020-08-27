VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'add data enviroment to the program and left it to default name
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open ("studenttable")
End Sub
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox ("enter values")
        Text1.SetFocus
    End If
    
    On Error GoTo err
        cn.Execute ("insert into student values('" & Text1.Text & "')")
        MsgBox ("value inserted")
        Text1.Text = ""
        Text1.SetFocus
err:
    MsgBox ("record already exits")
    Text1.Text = ""
    Text1.SetFocus
        
End Sub


