VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   15150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "My programs"
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "edit"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   7215
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "Programs list"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Code"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Language"
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "Setting"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
