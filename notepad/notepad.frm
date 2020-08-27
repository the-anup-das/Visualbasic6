VERSION 5.00
Begin VB.Form notepad 
   Caption         =   "Notepad"
   ClientHeight    =   4380
   ClientLeft      =   1680
   ClientTop       =   1710
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9255
   Begin VB.TextBox Text1 
      Height          =   4215
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu submnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu submnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu submnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu submnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu submnucut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu submnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu submnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu submnufind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu submnureplace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu submnugoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
      Begin VB.Menu submnuselectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu submnutimedate 
         Caption         =   "&Time /Date"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "F&ormat"
      Begin VB.Menu submnuwordwarp 
         Caption         =   "&Word Warp"
         Shortcut        =   ^M
      End
      Begin VB.Menu submnufont 
         Caption         =   "&Font"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu submnucursor 
         Caption         =   "&Cursor Status "
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu submnuabout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "notepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dc As Boolean
Public st As String, f As String, flen As Integer, fn As Boolean, c As Integer





Private Sub submnuabout_Click()
    aboutme.Show
End Sub

Private Sub submnunew_Click()
        Text1.Text = ""
End Sub

Private Sub submnuexit_Click()
    End
End Sub

Private Sub submnucut_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText, vbCFText
    Text1.SelText = " "
End Sub

Private Sub submnucopy_Click()
    Clipboard.Clear
       Clipboard.SetText Text1.SelText, vbCFText
End Sub

Private Sub submnupaste_Click()
    If Clipboard.GetFormat(vbCFText) Then
        Text1.Text = Text1.Text + Clipboard.GetText()
    Else
        MsgBox "the clipboard is empty "
    End If
End Sub

Private Sub submnufind_Click()
            Find.Show
End Sub

Private Sub submnureplace_Click()
            replace.Show
End Sub

Private Sub submnuselectall_Click()
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub submnutimedate_Click()
        If dc Then
         Text1.Text = Time() + Date
         dc = False
    Else
        Text1.Text = ""
       Text1.Text = Time() + Date
    End If
End Sub

Private Sub submnuwordwarp_Click()
    submnuwordwarp.Checked = True
End Sub
Private Sub Text1_Change()
        dc = True
End Sub
