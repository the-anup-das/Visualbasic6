VERSION 5.00
Begin VB.Form count 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Frequency of Character"
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
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox resulttxt 
      Height          =   975
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txt 
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Your Text"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "count"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st As String, length As Integer
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer, h As Integer, i As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer, o As Integer, p As Integer, q As Integer, r As Integer, s As Integer, t As Integer, u As Integer, v As Integer, w As Integer, x As Integer, y As Integer, z As Integer
Dim co, msg As Integer
Private Sub Command1_Click()
st = txt.Text
length = Len(st)
        For co = 1 To length Step 1
            Select Case Mid(st, co, 1)
                Case "a":
                    a = a + 1
                Case "b":
                    b = b + 1
                Case "c":
                    c = c + 1
                Case "d":
                    d = d + 1
                 Case "e":
                    e = e + 1
                 Case "f":
                    f = f + 1
                 Case "g":
                    g = g + 1
                 Case "h":
                    h = h + 1
                Case "i":
                    i = i + 1
                Case "j":
                    j = j + 1
                Case "k":
                    k = k + 1
                Case "l":
                    l = l + 1
                Case "m":
                    m = m + 1
                Case "n":
                    n = n + 1
                Case "o":
                    o = o + 1
                Case "p":
                    p = p + 1
                Case "q":
                    q = q + 1
                Case "r":
                    r = r + 1
                Case "s":
                    s = s + 1
                Case "t":
                    t = t + 1
                Case "u":
                    u = u + 1
                Case "v":
                    v = v + 1
                Case "w":
                    w = w + 1
                Case "x":
                    x = x + 1
                Case "y":
                        y = y + 1
                Case "z"
                    z = z + 1
                Case Else:
                    msg = MsgBox("No character Inserted", vbOKOnly + vbCritical, "Error")
                End Select
                'resulttxt.Text = "Character's are Found " + vbNewLine
                
        Next co
            Call display("a", a)
            Call display("b", b)
            Call display("c", c)
            Call display("d", d)
            Call display("e", e)
            Call display("f", f)
            Call display("e", e)
            Call display("g", g)
            Call display("h", h)
            Call display("i", i)
            Call display("j", j)
            Call display("k", k)
            Call display("l", l)
            Call display("m", m)
            Call display("n", n)
            Call display("o", o)
            Call display("p", p)
            Call display("q", q)
            Call display("r", r)
            Call display("s", s)
            Call display("t", t)
            Call display("u", u)
            Call display("v", v)
            Call display("w", w)
            Call display("x", x)
            Call display("y", y)
            Call display("z", z)
End Sub
Sub display(ByVal ch As String, ByVal tym As Integer)
    resulttxt.Text = resulttxt.Text + ch + "=" + str(tym)
End Sub
