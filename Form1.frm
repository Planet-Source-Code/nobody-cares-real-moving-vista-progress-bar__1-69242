VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   600
      TabIndex        =   3
      Text            =   "50"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   600
      TabIndex        =   2
      Text            =   "100"
      Top             =   1065
      Width           =   3375
   End
   Begin Project1.VistaProgress VP 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      _ExtentX        =   9551
      _ExtentY        =   661
      Value           =   50
   End
   Begin VB.Label Label1 
      Caption         =   "Max                                 Value"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   30
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
VP.Width = Me.ScaleWidth
End Sub

Private Sub Text1_Change()
VP.Max = Text1.Text
End Sub

Private Sub Text2_Change()
VP.Value = Text2.Text
End Sub
