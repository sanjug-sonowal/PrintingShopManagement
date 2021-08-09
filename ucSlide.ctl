VERSION 5.00
Begin VB.UserControl ucSlide 
   BackColor       =   &H00D6E4FC&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin Proyecto1.ucNeumorphism ThumbBack 
      Height          =   660
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      Invert          =   -1  'True
      Distance        =   3
      Radius          =   1000
      Intencity       =   50
      Blur            =   3
      ShadowColor     =   15652797
      BackColor       =   16379364
      Enabled         =   0   'False
      MousePointer    =   0
      MouseToParent   =   -1  'True
      ButtonStyle     =   0   'False
   End
   Begin Proyecto1.LabelPlus ThumbFore 
      Height          =   210
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      BackColor       =   16756370
      BackShadow      =   0   'False
      BorderCornerLeftTop=   1000
      BorderCornerRightTop=   1000
      BorderCornerBottomRight=   1000
      BorderCornerBottomLeft=   1000
      Caption         =   "ucSlide.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      MouseToParent   =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin Proyecto1.ucNeumorphism ThumbFore1 
      Height          =   570
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      Invert          =   -1  'True
      Distance        =   3
      Radius          =   1000
      Intencity       =   50
      Blur            =   3
      StatePressed    =   -1  'True
      ShadowColor     =   15652797
      BackColor       =   16379364
      Enabled         =   0   'False
      MousePointer    =   0
      MouseToParent   =   -1  'True
      ButtonStyle     =   0   'False
   End
   Begin Proyecto1.LabelPlus TrackFore 
      Height          =   105
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   185
      BackColor       =   16756370
      BackShadow      =   0   'False
      BorderCornerLeftTop=   100
      BorderCornerRightTop=   100
      BorderCornerBottomRight=   100
      BorderCornerBottomLeft=   100
      Caption         =   "ucSlide.ctx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      MouseToParent   =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin Proyecto1.ucNeumorphism TrackBack 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      Invert          =   -1  'True
      Distance        =   2
      Blur            =   2
      StatePressed    =   -1  'True
      BackColor       =   16313571
      MousePointer    =   0
      MouseToParent   =   -1  'True
      ButtonStyle     =   0   'False
   End
End
Attribute VB_Name = "ucSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Event Change(ByVal Value As Long)
Dim m_Percent As Long
Dim lDisplas  As Long
Dim nScale As Single
Dim m_Horizontal As Boolean
Dim m_Max As Double
Dim m_Min As Double
Dim m_Value As Double
Dim bFocus As Boolean

Public Property Let percent(ByVal NewValue As Double)
    m_Percent = NewValue
    'PropertyChanged "Percent"
End Property

Public Property Get percent() As Double
    percent = m_Percent
End Property

Public Property Let min(ByVal NewValue As Double)
    m_Min = NewValue
    PropertyChanged "Min"
End Property

Public Property Get min() As Double
    min = m_Min
End Property

Public Property Let Max(ByVal NewValue As Double)
    m_Max = NewValue
    PropertyChanged "Max"
End Property

Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Horizontal(ByVal NewValue As Boolean)
    Dim xSize As Single
    If m_Horizontal = NewValue Then Exit Property
    m_Horizontal = NewValue
    If NewValue Then
        xSize = UserControl.Extender.Height
        UserControl.Extender.Height = UserControl.Extender.Width
        UserControl.Extender.Width = xSize
    Else
        xSize = UserControl.Extender.Width
        UserControl.Extender.Width = UserControl.Extender.Height
        UserControl.Extender.Height = xSize
    End If
    ThumbBack.Move 0, 0
    UserControl_Resize
    PropertyChanged "Horizontal"
End Property

Public Property Get Horizontal() As Boolean
    Horizontal = m_Horizontal
End Property

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(new_value As Double)
    Dim W As Long, H As Long

    m_Percent = Abs((Abs(new_value) - Abs(m_Min)) * 100 / (m_Max - m_Min))
    If m_Horizontal Then
        ThumbBack.Left = ((UserControl.ScaleWidth - ThumbBack.Width) * m_Percent / 100)
        W = ThumbBack.Left + ThumbBack.Width / 2 - TrackFore.Left
        If W <= 0 Then W = 1
        TrackFore.Width = W
    Else
        ThumbBack.Top = ((UserControl.ScaleHeight - ThumbBack.Height) * m_Percent / 100)
        H = UserControl.ScaleHeight - ThumbBack.Top + ThumbBack.Height / 2 - ThumbBack.Height
        If H <= 0 Then H = 1
        TrackFore.Height = H
        TrackFore.Top = UserControl.ScaleHeight - ThumbBack.Height - TrackFore.Height
        UserControl_Resize
    End If

    m_Value = new_value
    PropertyChanged "Value"
End Property

'Private Sub TrackBack_PrePaint(hdc As Long, x As Long, y As Long)
'    Dim i As Long, H As Long
'    H = TrackBack.Height - lDisplas * 2
'    For i = 0 To H Step H / 12
'        ThumbFore.DrawLine hdc, 0, i + lDisplas, 2, i + lDisplas, RGB(127, 127, 127)
'    Next
'End Sub

Private Sub UserControl_InitProperties()
    m_Horizontal = True
    m_Max = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Horizontal = .ReadProperty("Horizontal", True)
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 100)
        m_Value = .ReadProperty("Value", 0)
    End With
    Me.Value = m_Value
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Horizontal", m_Horizontal, True)
        Call .WriteProperty("Min", m_Min, 0)
        Call .WriteProperty("Max", m_Max, 100)
        Call .WriteProperty("Value", m_Value, 0)
    End With
End Sub

Private Sub ThumbBack_PostPaint(ByVal hdc As Long)
    Dim x As Long, y As Long
    x = (ThumbBack.Width / 2 - ThumbFore1.Width / 2) + (ThumbFore1.Distance + ThumbFore1.Blur * 2) * nScale
    y = (ThumbBack.Height / 2 - ThumbFore1.Height / 2) + (ThumbFore1.Distance + ThumbFore1.Blur * 2) * nScale
    
    ThumbFore1.Draw hdc, x, y
    ThumbFore.Draw hdc, 0, ThumbBack.Width / 2 - ThumbFore.Width / 2, ThumbBack.Height / 2 - ThumbFore.Height / 2
End Sub

Public Sub FillColor(oColor As OLE_COLOR)
    UserControl.BackColor = oColor
    UserControl.BackStyle = 1
   
End Sub


Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Paint()
    Dim tRect As RECT
    If bFocus Then
        tRect.Right = UserControl.ScaleWidth
        tRect.Bottom = UserControl.ScaleHeight
        DrawFocusRect UserControl.hdc, tRect
    End If
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim W As Long, H As Long, lPercent As Long
    
    If Button = vbLeftButton Then
        If m_Horizontal Then
            If x > ThumbBack.Width / 2 Then
                If x < UserControl.ScaleWidth - ThumbBack.Width / 2 Then
                    ThumbBack.Left = x - ThumbBack.Width / 2
                Else
                    ThumbBack.Left = UserControl.ScaleWidth - ThumbBack.Width
                End If
            Else
                ThumbBack.Left = 0
            End If
            
            W = ThumbBack.Left + ThumbBack.Width / 2 - TrackFore.Left
            If W > 0 Then
                TrackFore.Width = W
                lPercent = CLng(W * 100 / (UserControl.ScaleWidth - ThumbBack.Width))
            Else
                TrackFore.Width = 1
                lPercent = 0
            End If

        Else
            If y > ThumbBack.Height / 2 Then
                If y < UserControl.ScaleHeight - ThumbBack.Height / 2 Then
                    ThumbBack.Top = y - ThumbBack.Height / 2
                Else
                    ThumbBack.Top = UserControl.ScaleHeight - ThumbBack.Height
                End If
            Else
                ThumbBack.Top = 0
            End If
            
            H = UserControl.ScaleHeight - ThumbBack.Height - ThumbBack.Top
            If H > 0 Then
                TrackFore.Top = UserControl.ScaleHeight - ThumbBack.Height / 2 - H + 1 * nScale
                TrackFore.Height = H
                lPercent = CLng(H * 100 / (UserControl.ScaleHeight - ThumbBack.Height))
            Else
                TrackFore.Top = UserControl.ScaleHeight - ThumbBack.Height / 2 + 1 * nScale
                TrackFore.Height = 1
                lPercent = 0
            End If
        
        End If
        If m_Percent <> lPercent Then
            
            m_Percent = lPercent

            m_Value = ((m_Max - m_Min) * m_Percent / 100) + m_Min
            RaiseEvent Change(lPercent)
        End If
    End If
End Sub



Private Sub UserControl_Resize()
    On Error Resume Next
    Dim L As Long, t As Long, W As Long, H As Long
    If m_Horizontal Then
    
        UserControl.Height = ThumbBack.Height * Screen.TwipsPerPixelY
        
        L = ThumbBack.Width / 2 - lDisplas
        t = UserControl.ScaleHeight / 2 - TrackBack.Height / 2
        W = UserControl.ScaleWidth - ThumbBack.Width + lDisplas * 2
        H = ThumbBack.Height / 2
        
        TrackBack.Move L, t, W, H

        ThumbBack.Left = ((UserControl.ScaleWidth - ThumbBack.Width) * m_Percent / 100)  '- ThumbBack.Width / 2
        
        L = TrackBack.Left + lDisplas '- 2 * nScale
        
        W = ThumbBack.Left + ThumbBack.Width / 2 - TrackFore.Left
        H = TrackBack.Height - lDisplas * 2
        t = UserControl.ScaleHeight / 2 - H / 2
        If W < 0 Then W = 1
        
        TrackFore.Move L, t, W, H
    Else
    
        UserControl.Width = ThumbBack.Width * Screen.TwipsPerPixelX
        W = ThumbBack.Width / 2
        L = UserControl.ScaleWidth / 2 - W / 2
        t = ThumbBack.Height / 2 - lDisplas
        
        H = UserControl.ScaleHeight - ThumbBack.Height + lDisplas * 2
        TrackBack.Move L, t, W, H
        
        H = (UserControl.ScaleHeight - ThumbBack.Height)
        ThumbBack.Top = H - (H * m_Percent / 100) '- ThumbBack.Height / 2
        'H = ThumbBack.Top + ThumbBack.Height / 2 - TrackFore.Top
        
        t = TrackBack.Top + lDisplas '- 1 * nScale
        W = TrackBack.Width - lDisplas * 2
        L = UserControl.ScaleWidth / 2 - W / 2
        H = UserControl.ScaleHeight - ThumbBack.Height - ThumbBack.Top
        If H < 0 Then H = 1
        t = UserControl.ScaleHeight - ThumbBack.Height / 2 - H '+ 1 * nScale
        'TrackFore.Top =UserControl.ScaleHeight - ThumbBack.Height / 2 - H + 1 * nScale
        
        TrackFore.Move L, t, W, H
    End If
End Sub

Private Sub UserControl_Show()
    nScale = ThumbBack.GetWindowsDPI
    lDisplas = (TrackBack.Distance + TrackBack.Blur * 2 + 1) * nScale
    
    If m_Horizontal Then
        ThumbBack.Top = 0
        'ThumbBack.Left = ((UserControl.ScaleWidth - ThumbBack.Width) * m_Percent / 100)
        'TrackBack.Height = ThumbBack.Height / 2
        'TrackFore.Height = TrackBack.Height - lDisplas * 2 + 2 * nScale
    Else
        
        'ThumbBack.Top = (UserControl.ScaleHeight - ThumbBack.Height) - ((UserControl.ScaleHeight - ThumbBack.Height) * m_Percent / 100)
        ThumbBack.Left = 0
        'TrackBack.Width = ThumbBack.Width / 2
        'TrackFore.Width = TrackBack.Width - lDisplas * 2 + 2 * nScale
    End If
    UserControl_Resize
End Sub

Private Sub UserControl_EnterFocus()
    bFocus = True
    UserControl.Refresh
End Sub

Private Sub UserControl_ExitFocus()
    bFocus = True
    UserControl.Refresh
End Sub
