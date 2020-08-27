VERSION 5.00
Begin VB.Form Find 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2040
   ClientLeft      =   2970
   ClientTop       =   3825
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6765
   Begin VB.CheckBox match 
      Caption         =   "Match Case"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton pos 
      Caption         =   "Find Next"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox findtxt 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim tym As Integer
Private Sub pos_Click()
                                   
                                    Dim arr(100) As Integer
                                    st = notepad.Text1.Text
                                    f = findtxt.Text
                                    flen = Len(f)
                                    s = Len(st)

                                    If match.Value = 0 Then
                                             st = UCase(st)
                                             f = UCase(f)
                                    End If
                                    
                                    c = 0
                                    For i = 1 To s Step 1
                                        If Mid(st, i, 1) = f Then
                                            arr(c) = i
                                            c = c + 1
                                        End If
                                        
                                    Next i
                           
                    If c <> 0 Then
                    
                            notepad.Text1.SelStart = arr(tym) - 1
                            notepad.Text1.SelLength = flen
                            tym = tym + 1
                             If tym = c Then
                                tym = 0
                            End If
                    Else
                            g = MsgBox("Word is not found", vbOKOnly + vbCritical, "Try Again")
                    End If
                                   
End Sub
