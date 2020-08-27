VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      DragMode        =   1  'Automatic
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "label"
      DragMode        =   1  'Automatic
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    If TypeOf Source Is TextBox Then
        Label1.Caption = Text1.Text
        End If
End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    Dim imgname
    
        If TypeOf Source Is TextBox Then
            imgname = Source.Text
        Else
            imgname = Source.Caption
        End If
        
        On Error GoTo NOIMAGE
            Picture1.Picture = LoadPicture(imgname)
            Exit Sub
NOIMAGE:
        MsgBox ("This is not a valid file name")
        
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
    If TypeOf Source Is Label Then
        Text1.Text = Label1.Caption
    End If
End Sub
