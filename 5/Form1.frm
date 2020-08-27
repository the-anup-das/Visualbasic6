VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton show 
      Caption         =   "show result"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox result 
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox st 
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your string "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
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
Private Sub show_Click()
s = st.Text
l = Len(s)
If l = 0 Then
result.Text = "enter some value "
st.SetFocus
Else
    result.Text = ""
    c = 0
    For i = 1 To l Step 1
        If Mid(s, i, 1) = " " Then
            Exit For 'to terminate for loop
        End If
        c = c + 1
    Next i
    result.Text = Str(c)
End If
End Sub
