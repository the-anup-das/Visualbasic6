VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form check 
   Caption         =   "check a string is palindrome or not"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show Previous inserted strings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   7
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Palindrome"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   9120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      DataField       =   "nonpalindrome"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   9120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox result 
      Height          =   855
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   735
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   5175
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   11280
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "DSN=stringtype"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "stringtype"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "nonpalindrome"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11160
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "DSN=stringtype"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "stringtype"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "palindrome"
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
   Begin VB.Label Label2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your string"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
            Dim str As String
            str = txt.Text
            s = UCase(str)
            
            If s = StrReverse(s) Then
                result.Text = "true"
                Text1.Text = str
                Adodc1.Recordset.AddNew
            Else
                result.Text = "false"
                Text3.Text = str
                Adodc2.Recordset.AddNew
            End If
            
End Sub

Private Sub Command2_Click()
    Histroy.Show
End Sub
