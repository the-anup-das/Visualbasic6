VERSION 5.00
Begin VB.Form show 
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Save your Contacts "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset



Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open ("phcn")
End Sub

Private Sub Command1_Click()
    If Text1.Text = "" And Text2.Text = "" Then
        uff = MsgBox("enter values", vbInformation)
        Text1.SetFocus
    End If
    
    On Error GoTo err
        If Len(Text2.Text) = 10 Then
        cn.Execute ("insert into cndetail values('" & Text1.Text & "','" & Text2.Text & "')")
        MsgBox ("value inserted")
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
        
        Else
            MsgBox ("Enter valid mobile number")
            Text2.Text = ""
            Text2.SetFocus
        End If
err:
            MsgBox ("record already exits")
            Text1.Text = ""
            Text2.Text = ""
            Text1.SetFocus
End Sub

