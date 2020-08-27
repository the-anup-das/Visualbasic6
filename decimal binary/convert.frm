VERSION 5.00
Begin VB.Form convert 
   Caption         =   "Binary - Decimal"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox decimaltxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox binary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton d2b 
      Caption         =   "Decimal to Binary"
      BeginProperty Font 
         Name            =   "Cracked Johnnie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton b2d 
      Caption         =   "Binary to Decimal"
      BeginProperty Font 
         Name            =   "Cracked Johnnie"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Decimal"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Binary"
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim c As Boolean, d As Boolean
 Private Sub b2d_Click()
            Dim arr(100), r As Integer
            b = Val(binary.Text)
            If b = Val(binary.Text) Then
                c = True
            End If
            i = 1
            res = 0
            While (b <> 0)
                r = b Mod 10
                res = res + r * i
                 i = i * 2
                b = b \ 10
                
            Wend
            
            If c = True Then
                 decimaltxt.Text = ""
                decimaltxt.Text = decimaltxt.Text + Str(res)
                c = False
            Else
                decimaltxt.Text = ""
            End If
End Sub



Private Sub d2b_Click()
            Dim arr(100) As Integer
            b = Val(decimaltxt.Text)
            
                If b = Val(decimaltxt.Text) Then
                    d = True
                End If
            
                          i = 0
                          res = 0
                    
                          While (b <> 0)
                              r = b Mod 2
                              arr(i) = r
                              b = b \ 2
                              i = i + 1
                              
                          Wend
                          
                              If d = True Then
                                    binary.Text = ""
                                     j = i - 1
                                     While (j >= 0)
                                        binary.Text = binary.Text + Str(arr(j))
                                         j = j - 1
                                     Wend
                                     d = False
                             Else
                                     binary.Text = ""
                            End If
End Sub
