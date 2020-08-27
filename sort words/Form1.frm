VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Enter"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str1 As String, tmp As String
Dim n As Integer, i As Integer, j As Integer



Private Sub Command2_Click()
    List1.Clear
    On Error GoTo Error
    n = CInt(InputBox("Enter how many Strings you want to enter"))
    
 For i = 0 To n - 1
        str1 = InputBox("Enter your string " & Str(i))
        List1.AddItem str1, i
    Next i
    Exit Sub
Error:
    MsgBox "Enter Some value"
End Sub

Private Sub Command1_Click()
    For i = 0 To List1.ListCount - 1
        For j = 1 To (List1.ListCount - 1) - i
            If StrComp(List1.List(j), List1.List(j - 1), vbTextCompare) < 0 Then
                tmp = List1.List(j - 1)
                List1.List(j - 1) = List1.List(j)
                List1.List(j) = tmp
            End If
        Next j
    Next i
    
End Sub
