VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "New Fader Control"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin Project1.Fader Fader1 
      Height          =   3015
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5318
   End
   Begin Project1.Fader Fader1 
      Height          =   1695
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2990
   End
   Begin Project1.Fader Fader1 
      Height          =   3015
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5318
   End
   Begin Project1.Fader Fader1 
      Height          =   1695
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2990
   End
   Begin Project1.Fader Fader1 
      Height          =   1215
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2143
   End
   Begin Project1.Fader Fader1 
      Height          =   3015
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5318
   End
   Begin Project1.Fader Fader1 
      Height          =   1215
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2143
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserControl11_Scrolling()
    Me.Caption = UserControl11.Value
End Sub

Private Sub UserControl12_Scrolling()
Me.Caption = UserControl12.Value

End Sub
