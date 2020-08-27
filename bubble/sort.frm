VERSION 5.00
Begin VB.Form sort 
   Caption         =   "Bubble Sort"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox inserttxt 
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton decending 
      Caption         =   "Decending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton assending 
      Caption         =   "Assending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox result 
      Height          =   1095
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox sizetxt 
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Your inserted array :"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
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
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "How many number you want in array :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(20) As Integer, size As Integer, c As Integer, i As Integer
Sub insert()
size = Val(sizetxt.Text)
    
    If c <> 0 Then
        i = MsgBox("We have values in array!! u want to use those values", vbYesNo, "Conformation")
        If i = 7 Then
            For i = 1 To size Step 1
                arr(i) = InputBox("Enter your Numbers (After inserting a number press enter to other number) :", "Input")
                inserttxt.Text = inserttxt.Text + Str(arr(i))
            Next i
        End If
    End If
    
    If c = 0 Then
            For i = 1 To size Step 1
                arr(i) = InputBox("Enter your Numbers (After inserting a number press enter to other number) :", "Input")
                inserttxt.Text = inserttxt.Text + Str(arr(i))
            Next i
        c = c + 1
    End If
End Sub
Sub out()
    result.Text = ""
    For i = 1 To size Step 1
    result.Text = result.Text + Str(arr(i))
    Next i
End Sub
Private Sub assending_Click()

    insert
    For i = 1 To size Step 1
        For j = 1 To size - i Step 1
            If arr(j) >= arr(j + 1) Then
                small = arr(j + 1)
                arr(j + 1) = arr(j)
                arr(j) = small
            End If
        Next j
    Next i
    out
    
End Sub


Private Sub decending_Click()
        insert
        For i = 1 To size Step 1
            For j = 1 To size - i Step 1
                If arr(j) <= arr(j + 1) Then
                    small = arr(j + 1)
                    arr(j + 1) = arr(j)
                    arr(j) = small
                End If
            Next j
        Next i
        out
End Sub
