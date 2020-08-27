VERSION 5.00
Begin VB.Form expense 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "refresh"
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox lefttxt 
      DataField       =   "left"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   855
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox costtxt 
      DataField       =   "cost"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   3840
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox expensetxt 
      DataField       =   "exlist"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add "
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label labeltime 
      DataField       =   "time"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label dlable 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   7200
      X2              =   7200
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   9000
      X2              =   9000
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   9000
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   9000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   3720
      X2              =   3720
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   5520
      X2              =   5520
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "Left ( Money)"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Cost"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Expense"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Date"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "expense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim temprs As ADODB.Recordset
Dim totalmoney, m As Integer, reply As Integer
Dim autocal As Boolean
Public Sub Command1_Click()
    On Error GoTo cancelupdate
       
       '     Call clr
        '        m = Adodc1.Recordset.Fields("left").Value
       'lefttxt.Text = m
       'dlable.Caption = Date
       'labeltime.Caption = Time
       cn.Execute ("insert into expanses values ('" & dlable.Caption & "','" & expensetxt.Text & "','" & costtxt.Text & "','" & lefttxt.Text & "','" & labeltime.Caption & "')")
       MsgBox "Inserted"
       Call clr
       
         Set temprs = cn.Execute("select left from expanses")
         While Not temprs.EOF
            m = temprs(0)
            temprs.MoveNext
          Wend
          
       lefttxt.Text = m
       dlable.Caption = Date
       labeltime.Caption = Time
       
       Exit Sub

cancelupdate:
   MsgBox Err.Description
   Call clr
   dlable.Caption = Date
   labeltime.Caption = Time
End Sub



Private Sub costtxt_GotFocus()
    autocal = True                      'check any vales insert in costtxt
End Sub

Private Sub costtxt_Change()
    If autocal = True Then
         lefttxt.Text = m - Val(costtxt.Text)   'calculate left money
            If Val(lefttxt.Text) < 0 Then
                reply = MsgBox("Want to Add Money?", vbYesNoCancel + vbExclamation, "Incificent Balence")
                'Command1.Enabled = False
            Else
                Command1.Enabled = True
            End If
            
            Select Case reply
                Case 6:
                    moneygot.Show
                    Command1.Enabled = True
                Case 7:
                    Command1.Enabled = False
                Case 2:
                    Command1.Enabled = False
            End Select
    End If
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set temprs = New ADODB.Recordset
    cn.Open ("pmdsn")
    
     Set rs = cn.Execute("select *from expanses")
     Set temprs = cn.Execute("select left from expanses")
    
      If rs.BOF Then
        m = moneygot.Adodc1.Recordset.Fields("amount").Value
        lefttxt.Text = m
     Else
         
         While Not temprs.EOF
            m = temprs(0)
            temprs.MoveNext
          Wend
        lefttxt.Text = m
     End If
    
    dlable.Caption = Date
    labeltime.Caption = Time
    Call clr
  
End Sub
Private Sub Command4_Click()
    End
End Sub

Private Sub clr()

expensetxt.Text = ""
costtxt.Text = ""

End Sub

