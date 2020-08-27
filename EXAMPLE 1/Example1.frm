VERSION 5.00
Begin VB.Form Example1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Modem"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   2240
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "CD-ROM"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1720
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "RAM"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Sound Card"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Caption         =   "VISA"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "BOI"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "SBI"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "American Express"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Mastercard"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Select Option Items"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Form of Payment "
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
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Example1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

