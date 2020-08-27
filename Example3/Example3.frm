VERSION 5.00
Begin VB.Form Example3 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option6 
      Caption         =   "Chorome "
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Firefox"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Internet Explorer"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Other"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select a Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a Language"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
      Begin VB.OptionButton Option1 
         Caption         =   "Visual Basic"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Java"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   990
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "c"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1740
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "c++"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   2490
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Other"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   3240
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "My Favorite Browser"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "My Favorite language "
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Example3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
