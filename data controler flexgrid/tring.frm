VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "tring.frx":0000
      Height          =   1575
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   2
      _Band(0)._MapCol(0)._Name=   "palin"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "no"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(1)._Alignment=   7
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   960
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=strtype2"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "strtype2"
      OtherAttributes =   ""
      UserName        =   "me"
      Password        =   "tiger"
      RecordSource    =   "Palin"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Himalaya"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim palind As String, i As Integer

Private Sub Command1_Click()
'palind = Adodc1.Recordset.Fields("Palin").Value ' to store recordset current value
palind = Adodc1.Recordset.RecordCount
MsgBox palind
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Fields("palin").Value = Text1.Text
Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
    For i = 1 To Adodc1.Recordset.RecordCount
        List1.AddItem Adodc1.Recordset.Fields("palin").Value
        Adodc1.Recordset.MoveNext
    Next i
   ' Do While Not (Adodc1.Recordset.EOF)
    '    List1.AddItem Adodc1.Recordset.Fields("palin").Value
    '    Adodc1.Recordset.MoveNext
   ' Loop
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Activate()
Adodc1.Refresh
End Sub

