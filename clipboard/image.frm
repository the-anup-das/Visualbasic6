VERSION 5.00
Begin VB.Form image 
   Caption         =   "Coping from clipboard"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Copy from clipboard"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy to Clipboard"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      Picture         =   "image.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetData Picture1.image, vbCFBitmap
End Sub

Private Sub Command2_Click()
If Clipboard.GetFormat(vbCFBitmap) Then
    Picture2.Picture = Clipboard.GetData()
Else
    MsgBox "the clipboard is empty "
End If

End Sub
