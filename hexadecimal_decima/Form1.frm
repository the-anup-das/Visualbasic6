VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox decimaltxt 
      Height          =   855
      Left            =   2520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox hexadecimaltxt 
      Height          =   855
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Hexadecimal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Boolean



Private Sub decimaltxt_Click()
 con = False
End Sub

Private Sub hexadecimaltxt_Click()
    con = True
End Sub

Private Sub show_Click()
If con = True Then
    Dim d, n, c As Integer
    Dim h As String
    
    h = hexadecimaltxt.Text
    n = Len(h)
    c = 0
    
    For i = n To 1 Step -1
        If Asc(Mid(h, i, 1)) >= Asc(0) And Asc(Mid(h, i, 1)) <= Asc(9) Then
            d = Mid(h, i, 1) - 0
            a = a + (16 ^ c) * d
            c = c + 1
            
        ElseIf Asc(Mid(h, i, 1)) >= Asc("A") And Asc(Mid(h, i, 1)) <= Asc("F") Then
            d = Asc(Mid(h, i, 1)) - 55
            a = a + (16 ^ c) * d
            c = c + 1
        End If
    Next i
    decimaltxt.Text = Str(a)
Else
    Dim arr(1000), dn As Integer
    dn = Val(decimaltxt.Text)
    c = 0
    length = 1
    
    While (dn <> 0)
        arr(c) = dn Mod 16
        dn = dn / 16
        c = c + 1
    Wend
        For i = c - 1 To 0 Step -1
        If arr(i) >= 0 And arr(i) <= 9 Then
              h = h & arr(i)
        ElseIf arr(i) >= 10 And arr(i) <= 15 Then
            d = arr(i) + 55
            h = h & Chr(d)
        End If
    Next i
    hexadecimaltxt.Text = h
End If
End Sub
