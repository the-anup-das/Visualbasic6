VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Marks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Roll:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim firsttime As Boolean


Private Sub Command6_Click()
    If firsttime = False Then
         Call Command5_Click
        firsttime = True
    Else
    If Not rs.EOF Then
        Text1.Text = rs(0)
        Text2.Text = rs(1)
        Text3.Text = rs(2)
        rs.MoveNext
    End If
End If
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open ("baby")
End Sub
Private Sub Command1_Click()
 On Error GoTo err
 If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
    MsgBox "Entry First"
 Else
    cn.Execute ("insert into table1 values(' " & Text1.Text & " ' , ' " & Text2.Text & " ' ,  ' " & Text3.Text & "')")
    MsgBox "Inseted table1"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End If
Exit Sub
err:
 MsgBox "Error"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End Sub
Private Sub Command2_Click()
Dim roll As Integer
On Error GoTo err
 If Text1.Text = "" Then
    MsgBox "Entry First"
 Else
    roll = Val(Text1.Text)
    cn.Execute ("delete from table1 where Roll =(" & roll & ")")
    MsgBox "Deleted table1"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End If
Exit Sub
err:
 MsgBox "Error"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End Sub
Private Sub Command3_Click()
Dim roll, marks As Integer
Dim name As String
  On Error GoTo err
 If Text1.Text = "" Then
    MsgBox "Entry First"
 Else
 roll = Val(Text1.Text)
 marks = Val(Text3.Text)
 
    cn.Execute ("update table1 set Marks=(' " & marks & " ') where Roll=(" & roll & ")")
    MsgBox "Update table1"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End If
Exit Sub
err:
 MsgBox "Error"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End Sub
Private Sub Command4_Click()
Dim roll As Integer
     On Error GoTo err
 If Text1.Text = "" Then
    MsgBox "Entry First"
 Else
 roll = Val(Text1.Text)
    Set rs = cn.Execute("select Name , Marks from table1 where Roll=(" & roll & ")")
    MsgBox "Selected table1"
    
        
        Text2.Text = rs(0)
        Text3.Text = rs(1)
        Text1.SetFocus
End If
Exit Sub
err:
 MsgBox "Error"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End Sub

Public Sub Command5_Click()
    Dim roll As Integer
     On Error GoTo err
 
    Set rs = cn.Execute("select *from table1")
        Exit Sub
err:
 MsgBox "Error"
End Sub

