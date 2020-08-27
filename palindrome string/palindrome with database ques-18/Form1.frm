VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form histroy 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton last 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton previous 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton next 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      Begin VB.OptionButton Option2 
         Caption         =   "Non-Palindrome"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Palindrome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton first 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   11400
      Top             =   13080
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.TextBox Text3 
      DataField       =   "nonpalindrome"
      DataSource      =   "Adodc2"
      Height          =   975
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      DataField       =   "Palindrome"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   18000
      Top             =   12360
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.Label Label1 
      Caption         =   "Select Entered String Type"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Histroy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p, n As Boolean

Private Sub first_Click()
    If p = True Then
        Adodc1.Recordset.MoveFirst
    End If
        
    If n = True Then
        Adodc2.Recordset.MoveFirst
    End If
End Sub

Private Sub last_Click()

    If p = True Then
        Adodc1.Recordset.MoveLast
    End If
        
    If n = True Then
        Adodc2.Recordset.MoveLast
    End If

End Sub

Private Sub next_Click()
If p = True Then
        Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then
                j = MsgBox("Last Record", vbOKOnly + vbInformation)
                Adodc1.Recordset.MoveLast
            End If
    End If
    
    If n = True Then
     Adodc2.Recordset.MoveNext
        If Adodc2.Recordset.EOF Then
            j = MsgBox("Last Record", vbOKOnly + vbInformation)
            Adodc2.Recordset.MoveLast
        End If
    End If
        
End Sub

Private Sub previous_Click()
    If p = True Then
        Adodc1.Recordset.MovePrevious
            If Adodc1.Recordset.BOF Then
               j = MsgBox("First Record", vbOKOnly + vbInformation)
                Adodc1.Recordset.MoveFirst
            End If
        End If
    If n = True Then
        Adodc2.Recordset.MovePrevious
            If Adodc2.Recordset.BOF Then
                j = MsgBox("First Record", vbOKOnly + vbInformation)
                Adodc2.Recordset.MoveFirst
            End If
    End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Text1.Visible = True
    Text3.Visible = False
    p = True
    n = False
End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Text1.Visible = False
        Text3.Visible = True
        n = True
        p = False
    End If
End Sub
