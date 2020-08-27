VERSION 5.00
Begin VB.Form frmCheck 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Even or Odd Numbers "
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Your End Value"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Strating value"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stv As Integer, ev As Integer
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
    stv = Val(Text1.Text)
    ev = Val(Text2.Text)
    For i = stv To ev
        If i Mod 2 = 0 Then
            'Text3.Text = Text3.Text + Str(i)
            On Error GoTo err
            cn.Execute ("insert into even values('" & i & "')")
        Else
            On Error GoTo err
            cn.Execute ("insert into odd values('" & i & "')")
        End If
    Next i
err:
    MsgBox "contains error" & vbCrLf & err.Description
    
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open ("numtypeo")
End Sub
