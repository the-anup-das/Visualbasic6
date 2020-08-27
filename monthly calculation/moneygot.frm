VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form moneygot 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "total"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   7920
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   2
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "FILE NAME=S:\VB\monthly calculation\pmdsn.dsn"
      OLEDBString     =   ""
      OLEDBFile       =   "S:\VB\monthly calculation\pmdsn.dsn"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "moneygot"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton svbttm 
      Caption         =   "Save"
      Height          =   735
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      DataField       =   "limit"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   5880
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "amount"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   9840
      X2              =   9840
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   5760
      X2              =   5760
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   3960
      X2              =   3960
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   2160
      X2              =   2160
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   7800
      X2              =   7800
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   360
      X2              =   9840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   360
      X2              =   9840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   360
      X2              =   9840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Limit"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Time"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      DataField       =   "time"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "moneygot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim autocal As Boolean
Dim m As Integer
Public Sub Form_Load()
Label1.Caption = Date
Label3.Caption = Time
Text1.Text = ""
Text2.Text = ""
m = Adodc1.Recordset.Fields(1).Value
Text3.Text = m
End Sub

Public Sub svbttm_Click()
    Adodc1.Recordset.AddNew
     Call Form_Load
End Sub

Private Sub Text1_GotFocus()
    autocal = True
End Sub
Private Sub Text1_Change()
    If autocal = True Then
        Text3.Text = m + Val(Text1.Text)
    End If
End Sub


