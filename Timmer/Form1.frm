VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   30
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Timer watchTimer 
         Interval        =   1000
         Left            =   1200
         Top             =   2520
      End
      Begin VB.CommandButton Command4 
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label secLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Mistral"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         TabIndex        =   3
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label minLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label hourLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000010&
         BackStyle       =   1  'Opaque
         Height          =   2175
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   2415
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub watchTimer_Timer()
Dim hur, min, sec As Integer
Dim abc As Date

hourLabel.Caption = Hour(Now)
minLabel.Caption = Minute(Now)
secLabel.Caption = Second(Now)
End Sub
