VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   12
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Display Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   7
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Enter The Emp_Id Whose Record You Want To Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Enter The Emp_Id Whose Record You Want To See"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Enter The Name, Emp_Id And Salary of The Employee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5895
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
    cn.Open ("Emp")
End Sub
Private Sub Command1_Click()
    
   cn.Execute ("insert into Employee Values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "')")
   MsgBox ("Data Inserted")
   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
   Text1.SetFocus
   
End Sub
Private Sub Command2_Click()

    Dim n As Integer
    n = Val(Text4.Text)
    Set rs = cn.Execute("Select Emp_name from Employee where Emp_Id=val('" & Str(n) & "')")
    Text5.Text = rs(0)
    
End Sub
Private Sub Command3_Click()

    Dim n As Integer
    n = Val(Text6.Text)
    Set rs = cn.Execute("Delete * from Employee where Emp_Id=val('" & Str(n) & "')")
    MsgBox ("Record Deleted")
    
End Sub
Private Sub Command4_Click()

    Set rs = cn.Execute("Update Employee set Emp_Salary=Emp_Salary * 1.1 where Emp_Salary>2000")
    MsgBox ("Record Updated")
End Sub
