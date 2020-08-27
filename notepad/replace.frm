VERSION 5.00
Begin VB.Form replace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace"
   ClientHeight    =   2880
   ClientLeft      =   3165
   ClientTop       =   3450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5880
   Begin VB.CommandButton reall 
      Caption         =   "Replace All"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton replacecmd 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox replacetxt 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox findtxt 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton pos 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox match 
      Caption         =   "Match Case"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Replace With"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1455
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
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "replace"
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
                                        If Mid(st, i, flen) = f Then
                                            arr(c) = i
                                            c = c + 1
                                        End If
                                        
                                    Next i
                           
                    If c <> 0 Then
                    
                            notepad.Text1.SelStart = arr(tym) - 1
                            notepad.Text1.SelLength = flen
                            
                            replacecmd.Enabled = True
                            
                            tym = tym + 1
                             If tym = c Then
                                tym = 0
                            End If
                    Else
                            g = MsgBox("Word is not found", vbOKOnly + vbCritical, "Try Again")
                    End If
                                   
End Sub


Private Sub replacecmd_Click()
    notepad.Text1.SelText = replacetxt.Text
    replacecmd.Enabled = False
End Sub

Private Sub reall_Click()
                                    Dim arr(100) As Integer
                                    st = notepad.Text1.Text
                                    f = findtxt.Text
                                    flen = Len(f)
                                    s = Len(st)

                                    If match.Value = 0 Then
                                             st = UCase(st)
                                             f = UCase(f)
                                    End If
                                    
                                    For i = 1 To s Step 1
                                        If Mid(st, i, flen) = f Then
                                            notepad.Text1.SelStart = i - 1
                                            notepad.Text1.SelLength = flen
                                            notepad.Text1.SelText = replacetxt.Text
                                        End If
                                        
                                    Next i
End Sub
