VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "ENTER THE EMPLOY ID YOU WANT TO DELETE"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Result"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Eploy Salary"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Employ Name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Employ id"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
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
    Set rs = New ADODB.Recordset
    cn.Open ("emp")
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
        k = MsgBox("Enter Some value", vbOKOnly, "Error")
    Else
        cn.Execute ("insert into employe values (' " & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')")
        MsgBox "Inserted"
    End If
    
End Sub

Private Sub Command2_Click()
   Set rs = cn.Execute("select emp_name from employe where emp_id=val('" & Text4.Text & "')")
    Text5.Text = rs(0)
End Sub

Private Sub Command3_Click()
    Set rs = cn.Execute("Delete  from employe  where emp_id=val('" & Text6.Text & "')")
    MsgBox "Deleted"
End Sub

Private Sub Command4_Click()

End Sub
