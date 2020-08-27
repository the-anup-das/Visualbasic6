VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String, l As Integer, ch As String, a(26) As Integer
Private Sub Form_Activate()
    str1 = InputBox("Enter Your Text")
    l = Len(str1)
    For i = 1 To 26
        a(i) = 0
    Next i
    
    For i = 1 To l Step 1
        ch = Mid(str1, i, 1)
        If Asc(ch) >= 96 And Asc(ch) <= 122 Then
            a(Asc(ch) - 96) = a(Asc(ch) - 96) + 1
        ElseIf Asc(ch) >= 64 And Asc(ch) <= 91 Then
            a(Asc(ch) - 64) = a(Asc(ch) - 64) + 1
        End If
    Next i
    For i = 1 To 26 Step 1
        Text1.Text = Text1.Text & Chr(i + 64) & vbTab & Str(a(i)) & vbNewLine
    Next i
End Sub

