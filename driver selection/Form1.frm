VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filFile 
      Height          =   1260
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   1440
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Text            =   "c:\"
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dirDirectory_Change()
    filFile.Path = dirDirectory.Path
End Sub


Private Sub drvDrive_Change()
    dirDirectory.Path = drvDrive.Drive
End Sub

Private Sub Form_Load()
drvDrive.Drive = "c:\"
dirDirectory.Path = "c:\"

filFile.Path = dirDirectory.Path
End Sub


Private Sub Text1_Change()
    If Right(Text1.Text, 1) = "\" Then
        drvDrive.Drive = Left(Text1.Text, 3)
        dirDirectory.Path = Text1.Text
    End If
        
End Sub
