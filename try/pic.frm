VERSION 5.00
Begin VB.Form pic 
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   Picture         =   "pic.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton preview 
      Caption         =   "Preview"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton next 
      Caption         =   "Next"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   6480
      Width           =   1935
   End
End
Attribute VB_Name = "pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer



Private Sub next_Click()
If i = 0 Then
    pic.Picture = LoadPicture("M:\wallpaper\pic\extra\1.jpeg")
    ElseIf i = 1 Then
        pic.Picture = LoadPicture("M:/wallpaper/pic/extra/3.jpg")
        ElseIf i = 2 Then
            pic.Picture = LoadPicture("M:\wallpaper\pic\extra\2.jpg")
            ElseIf i = 3 Then
                pic.Picture = LoadPicture("M:\wallpaper\pic\extra\4.jpg")
                ElseIf i = 4 Then
                    pic.Picture = LoadPicture("M:\wallpaper\pic\extra\5.jpg")
                    ElseIf i = 5 Then
                     pic.Picture = LoadPicture("M:\wallpaper\pic\extra\6.jpg")
                        ElseIf i = 6 Then
                         pic.Picture = LoadPicture("M:\wallpaper\pic\extra\7.jpg")
                            ElseIf i = 7 Then
                             pic.Picture = LoadPicture("M:\wallpaper\pic\extra\8.jpg")
End If
If i <= 7 Then
    i = i + 1
Else
    Call preview_Click
End If
End Sub

Private Sub preview_Click()
If i = 0 Then
    pic.Picture = LoadPicture("M:\wallpaper\pic\extra\1.jpeg")
    ElseIf i = 1 Then
        pic.Picture = LoadPicture("M:/wallpaper/pic/extra/3.jpg")
        ElseIf i = 2 Then
            pic.Picture = LoadPicture("M:\wallpaper\pic\extra\2.jpg")
            ElseIf i = 3 Then
                pic.Picture = LoadPicture("M:\wallpaper\pic\extra\4.jpg")
                ElseIf i = 4 Then
                    pic.Picture = LoadPicture("M:\wallpaper\pic\extra\5.jpg")
                    ElseIf i = 5 Then
                     pic.Picture = LoadPicture("M:\wallpaper\pic\extra\6.jpg")
                        ElseIf i = 6 Then
                         pic.Picture = LoadPicture("M:\wallpaper\pic\extra\7.jpg")
                            ElseIf i = 7 Then
                             pic.Picture = LoadPicture("M:\wallpaper\pic\extra\8.jpg")
End If
If i >= 0 Then
    i = i - 1
Else
    Call next_Click
End If
End Sub
