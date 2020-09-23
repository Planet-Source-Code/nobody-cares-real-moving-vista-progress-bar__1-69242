VERSION 5.00
Begin VB.UserControl VistaProgress 
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   3015
   ScaleWidth      =   7620
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   2160
   End
   Begin VB.PictureBox dr 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5400
      Picture         =   "VistaProgress.ctx":0000
      Top             =   0
      Width           =   30
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   240
      Picture         =   "VistaProgress.ctx":02BA
      Top             =   1680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image8 
      Height          =   330
      Left            =   10
      Stretch         =   -1  'True
      Top             =   10
      Width           =   5400
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "VistaProgress.ctx":05B1
      Top             =   0
      Width           =   30
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "VistaProgress.ctx":086A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   600
      Picture         =   "VistaProgress.ctx":0B05
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   345
      Left            =   1920
      Picture         =   "VistaProgress.ctx":0DAE
      Top             =   1680
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image7 
      Height          =   345
      Left            =   600
      Picture         =   "VistaProgress.ctx":10AF
      Top             =   1920
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "VistaProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'±âº» ¼Ó¼º °ª:
Const m_def_Max = 100
Const m_def_Value = 100
'¼Ó¼º º¯¼ö:
Dim m_Max As Long
Dim m_Value As Long


'°æ°í! ÁÖ¼®À¸·Î µÇ¾î ÀÖ´Â ´ÙÀ½ ÁÙÀº Á¦°ÅÇÏ°Å³ª ¼öÁ¤ÇÏÁö ¸¶½Ê½Ã¿À!
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
    Max = m_Max
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Property

'°æ°í! ÁÖ¼®À¸·Î µÇ¾î ÀÖ´Â ´ÙÀ½ ÁÙÀº Á¦°ÅÇÏ°Å³ª ¼öÁ¤ÇÏÁö ¸¶½Ê½Ã¿À!
'MemberInfo=8,0,0,100
Public Property Get Value() As Long
    Value = m_Value
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Property

'»ç¿ëÀÚ Á¤ÀÇ ÄÁÆ®·Ñ¿¡ ´ëÇÑ ¼Ó¼ºÀ» ÃÊ±âÈ­ÇÕ´Ï´Ù.
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Value = m_def_Value
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Sub

'ÀúÀå¼Ò¿¡¼­ ¼Ó¼º°ªÀ» ·ÎµåÇÕ´Ï´Ù.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Sub

'¼Ó¼º°ªÀ» ÀúÀå¼Ò¿¡ ±â·ÏÇÕ´Ï´Ù.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
On Error Resume Next
    Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Sub

Private Sub usercontrol_Resize()
Image2.Width = UserControl.Width - Image1.Width - Image3.Width
Image3.Left = Image2.Left + Image2.Width
On Error Resume Next
Image8.Width = UserControl.Width / m_Max * m_Value - 80
dr.Width = UserControl.Width / m_Max * m_Value
End Sub

Private Sub Timer1_Timer()
dr.Cls
'dr.PaintPicture Image1.Picture, 0, 0
'dr.PaintPicture Image2.Picture, Image1.Width, 0
'dr.PaintPicture Image3.Picture, Image2.Width + Image1.Width, 0

dr.PaintPicture Image5.Picture, 10, 15, Image2.Width + Image2.Left, Image2.Height - 10
dr.PaintPicture Image4.Picture, Image1.Width, 10
dr.PaintPicture Image6.Picture, UserControl.Width / m_Max * m_Value - Image6.Width, 10
dr.PaintPicture Image7.Picture, Image7.Left, 10
Image7.Left = Image7.Left + 100
If Image7.Left > UserControl.Width Then
Image7.Left = 0 - (Image7.Width * 3)
End If
Image8.Picture = dr.Image
End Sub

