VERSION 5.00
Begin VB.Form math 
   Caption         =   "nCr-nPr"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton npr 
      Caption         =   "nPr"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton ncr 
      Caption         =   "nCr"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox result 
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox rtxt 
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox ntxt 
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the Valu of r"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Value of  n "
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ncr_Click()
        Dim n, r, fn, fr, fd, s As Integer
        n = Val(ntxt.Text)
        r = Val(rtxt.Text)
        d = n - r
        
        fn = 1
        
        While n > 0
                fn = fn * n
                n = n - 1
        Wend
        
        fr = 1
        
        While r > 0
                fr = fr * r
                r = r - 1
        Wend
        
        fd = 1
        
        While d > 0
                fd = fd * d
                d = d - 1
        Wend
        
        s = fn \ (fr * fd)
        result.Text = Str(s)
End Sub

Private Sub npr_Click()
                    Dim n, r, fn, fd, sol As Integer
                            n = Val(ntxt.Text)
                            r = Val(rtxt.Text)
                            d = n - r
                            
                            fn = 1
                            
                            While n > 0
                                    fn = fn * n
                                    n = n - 1
                            Wend
                            
                            fd = 1
                            
                            While d > 0
                                    fd = fd * d
                                    d = d - 1
                            Wend
                            sol = fn \ fd
                            result.Text = Str(sol)
End Sub
