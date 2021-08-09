VERSION 5.00
Begin VB.UserControl ucList 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   840
   End
End
Attribute VB_Name = "ucList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As DEVMODE) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ButtonClick(ByVal Index As Integer)
Public Event ItemDblClick(ByVal Index As Integer)
Public Event ItemMouseUp(ByVal Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim m_ScrollMax As Long
Dim ItemCount As Long
Dim m_ScrollValue As Long
Dim OldhBmp As Long
Dim hDCMemory As Long
Dim hBmp As Long
Dim bMouseIn As Boolean
Dim mListIndex As Integer
Dim lDif As Long

Public Property Get Title(Index As Integer) As String
    Title = ucItem(Index).Title
End Property

Public Property Let Title(Index As Integer, NewValue As String)
    ucItem(Index).Title = NewValue
End Property

Public Property Get Artist(Index As Integer) As String
    Artist = ucItem(Index).Artist
End Property

Public Property Let Artist(Index As Integer, NewValue As String)
    ucItem(Index).Artist = NewValue
End Property

Public Property Get ListIndex() As Integer
    ListIndex = mListIndex
End Property

Public Property Let ListIndex(ByVal NewValue As Integer)
    mListIndex = NewValue
    ucnSelection.Top = ucItem(0).Height * mListIndex - m_ScrollValue - lDif
    If NewValue > -1 Then ucnSelection.Visible = True
    If ucnSelection.Top + ucnSelection.Height > UserControl.ScaleHeight Then
        ScrollValue = ucnSelection.Top - ucItem(0).Top
    ElseIf ucnSelection.Top < 0 Then
        ScrollValue = ucnSelection.Top
    End If
End Property

Public Property Let ButtonState(ByVal Index As Integer, ByVal isPlaying As Boolean)
    ucItem(Index).ButtonState = isPlaying
End Property

Public Property Get ButtonState(ByVal Index As Integer) As Boolean
    ButtonState = ucItem(Index).ButtonState
End Property

Public Property Get ItemPath(ByVal Index As Integer) As String
    ItemPath = CStr(ucItem(Index).Tag)
End Property

Public Property Let ItemPath(ByVal Index As Integer, ByVal NewValue As String)
    ucItem(Index).Tag = NewValue
End Property

Public Property Get ListCount() As Integer
    ListCount = ItemCount
End Property


Public Sub Add(Path As String, ByVal Title As String, ByVal Artist As String)
    Dim h As Long
    If ItemCount > 0 Then
        Load ucItem(ucItem.Count)
    End If

    With ucItem(ItemCount)
        .Tag = Path
        .SetData Title, Artist
        .ZOrder 0
        If ItemCount > 0 Then
            .Top = ucItem(ItemCount - 1).Top + ucItem(ItemCount - 1).Height
        End If
         m_ScrollMax = (.Height * ucItem.Count - UserControl.ScaleHeight)
        If m_ScrollMax < 0 Then
            .Visible = True
        End If
    End With
    
    ItemCount = ItemCount + 1


    '/ Screen.TwipsPerPixelY
    
    h = UserControl.ScaleHeight - (ucItem(0).Height * ucItem.Count - UserControl.ScaleHeight)
    If h < ucItem(0).Height Then h = ucItem(0).Height
    LpScroll.Height = h
    If m_ScrollMax > 0 Then
        h = UserControl.ScaleHeight - LpScroll.Height
        Me.ScrollValue = m_ScrollMax * (LpScroll.Top * 100 / h) / 100
    End If
End Sub

Public Sub DeleteItem(Index As Long)
    Dim i As Long, h As Long
    
    
    If Index < 0 Or Index > ItemCount Then Exit Sub

    If Index = ItemCount - 1 And ItemCount > 1 Then
    
        Unload ucItem(ItemCount - 1)
        If Me.ListIndex = ItemCount - 1 Then
            Me.ListIndex = Me.ListIndex - 1
        End If
    ElseIf ItemCount > 1 Then
        For i = Index To ItemCount - 2
            With ucItem(i)
                '.Artist = ucItem(i + 1).Artist
                '.Title = ucItem(i + 1).Title
                .SetData ucItem(i + 1).Title, ucItem(i + 1).Artist
                .Tag = ucItem(i + 1).Tag
                .ButtonState = ucItem(i + 1).ButtonState
            End With
        Next
        Unload ucItem(ItemCount - 1)
       
    End If
    
    ItemCount = ItemCount - 1
    
    If ItemCount = 0 Then
        ucItem(0).Visible = False
        ucnSelection.Visible = False
        Me.ListIndex = -1
    End If

    m_ScrollMax = (ucItem(0).Height * ucItem.Count - UserControl.ScaleHeight)
    
    h = UserControl.ScaleHeight - m_ScrollMax '(ucItem(0).Height * ucItem.Count - UserControl.ScaleHeight)
    If h < ucItem(0).Height Then h = ucItem(0).Height
    LpScroll.Height = h
    
    If m_ScrollMax > 0 Then
        h = UserControl.ScaleHeight - LpScroll.Height
        Me.ScrollValue = m_ScrollMax * (LpScroll.Top * 100 / h) / 100
    Else
        Me.ScrollValue = 0
    End If
    
End Sub

Public Property Get ScrollValue() As Long
    ScrollValue = m_ScrollValue
End Property

Public Property Let ScrollValue(new_value As Long)
    Dim i As Long
    
    'Call SendMessage(UserControl.ContainerHwnd, WM_SETREDRAW, 0&, 0&)
    If new_value < 0 Then
        new_value = 0
    ElseIf new_value > m_ScrollMax And m_ScrollMax > 0 Then
        new_value = m_ScrollMax + 1
    End If


   ' If m_ScrollValue = New_Value Then Exit Property
    m_ScrollValue = new_value
    For i = 0 To ItemCount - 1
        ucItem(i).Top = ucItem(i).Height * i - m_ScrollValue
        If ucItem(i).Top + ucItem(i).Height > 0 And ucItem(i).Top < UserControl.ScaleHeight Then
            ucItem(i).Visible = True
        Else
            ucItem(i).Visible = False
        End If
    Next
    LP1.Visible = m_ScrollValue > 0
    LP2.Visible = ucItem(ucItem.Count - 1).Top + ucItem(ucItem.Count - 1).Height > UserControl.ScaleHeight
    LP1.ZOrder
    LP2.ZOrder
    LpScroll.Top = (UserControl.ScaleHeight - LpScroll.Height) * (m_ScrollValue * 100 / m_ScrollMax) / 100
    ucnSelection.Top = ucItem(0).Height * mListIndex - m_ScrollValue - lDif
    ucnSelection.ZOrder 1
    'Call SendMessage(UserControl.ContainerHwnd, WM_SETREDRAW, 1&, 0&)
    'RedrawWindow UserControl.ContainerHwnd, ByVal &H0, 0, 1
End Property



Private Sub LP1_PostPaint(ByVal hdc As Long)
    Dim i As Long, lBF As Long
    Dim ucWidth As Long, Height As Long

    ucWidth = UserControl.ScaleWidth '/ Screen.TwipsPerPixelX
    Height = LP2.Height '/ Screen.TwipsPerPixelY

    For i = 0 To Height
        lBF = CLng(255 - 255 * (i / Height)) * &H10000
        Call AlphaBlend(hdc, 0, i, ucWidth, 1, hDCMemory, 0, i, ucWidth, 1, lBF)
    Next

End Sub

Private Sub LP2_PostPaint(ByVal hdc As Long)
    Dim i As Long, lBF As Long
    Dim ucWidth As Long, ucHeight As Long, Height As Long

    ucWidth = UserControl.ScaleWidth '/ Screen.TwipsPerPixelX
    ucHeight = UserControl.ScaleHeight '/ Screen.TwipsPerPixelY
 
    Height = LP2.Height '/ Screen.TwipsPerPixelY
 
    For i = 0 To Height
        lBF = CLng(255 * (i / Height)) * &H10000
        Call AlphaBlend(hdc, 0, i, ucWidth, 1, hDCMemory, 0, ucHeight - Height + i, ucWidth, 1, lBF)
    Next
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    bMouseIn = False
    Timer2.Interval = 50
End Sub

Private Sub Timer2_Timer()
    If bMouseIn Then
        If LpScroll.BackColorOpacity < 50 Then
            LpScroll.BackColorOpacity = LpScroll.BackColorOpacity + 5
        Else
            Timer2.Interval = 0
        End If
    Else
        If LpScroll.BackColorOpacity > 0 Then
            LpScroll.BackColorOpacity = LpScroll.BackColorOpacity - 5
        Else
            Timer2.Interval = 0
        End If
    End If
End Sub


Private Sub ucItem_ButtonClick(Index As Integer)
        RaiseEvent ButtonClick(Index)
        ucnSelection.Visible = True
End Sub

Private Sub ucItem_DblClick(Index As Integer)
    RaiseEvent ItemDblClick(Index)
End Sub

Private Sub ucItem_GotFocus(Index As Integer)
    UserControl.SetFocus
End Sub

Private Sub ucItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ucItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mListIndex = Index
    ucnSelection.Top = ucItem(0).Height * mListIndex - m_ScrollValue - lDif
    ucnSelection.Visible = True
End Sub

Private Sub ucItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_ScrollMax < 0 Then Exit Sub
    Timer1.Enabled = False
    Timer1.Interval = 1500
    Timer1.Enabled = True
    bMouseIn = True
    Timer2.Interval = 50
End Sub

Private Sub ucItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent ItemMouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_ScrollMax <= 0 Then
     RaiseEvent MouseDown(Button, Shift, x, y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim h As Long
    If Button = vbLeftButton And m_ScrollMax > 0 Then
        h = UserControl.ScaleHeight - LpScroll.Height
        If y < 0 Then y = 0
        If y = h Then y = h
        Me.ScrollValue = m_ScrollMax * (y * 100 / h) / 100
    End If
    ucItem_MouseMove 0, Button, Shift, 0, 0
End Sub

Private Sub UserControl_Paint()
    BitBlt hDCMemory, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc, 0, 0, vbSrcCopy

    'BitBlt hDCMemory, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, UserControl.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    CreateBuffer
    ucItem(0).Top = 0
    m_ScrollMax = -1
End Sub

Private Sub UserControl_Resize()
    Dim i As Long
    
    CreateBuffer
    For i = 0 To ucItem.Count - 1
        ucItem(i).Width = UserControl.ScaleWidth - LpScroll.Width * 2
    Next
    LP1.Top = -1 '-1 * Screen.TwipsPerPixelY
    LP1.Width = UserControl.ScaleWidth - LpScroll.Width
    LP2.Top = UserControl.ScaleHeight - LP2.Height + 1 '+ 1 * Screen.TwipsPerPixelY
    LP2.Width = UserControl.ScaleWidth - LpScroll.Width
    LpScroll.Left = UserControl.ScaleWidth - LpScroll.Width
    
    If lDif = 0 Then lDif = (ucnSelection.Distance + ucnSelection.Blur * 2) * ucnSelection.GetWindowsDPI
    ucnSelection.Move -lDif, -lDif, UserControl.ScaleWidth + lDif - 4 * ucnSelection.GetWindowsDPI, ucItem(0).Height + lDif * 2
End Sub

Public Sub Clear()
    Dim i As Long
    m_ScrollValue = -1
    m_ScrollMax = -1
    For i = ucItem.Count - 1 To 1 Step -1
        Unload ucItem(i)
    Next
    ucnSelection.Visible = False
    ucItem(0).Visible = False
    ucItem(0).ButtonState = False
    ItemCount = 0
End Sub



Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
End Sub

Private Sub CreateBuffer() 'Acrylic buffer
    Dim DC As Long
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
 
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
   ' hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY)
    ReleaseDC 0&, DC
    OldhBmp = SelectObject(hDCMemory, hBmp)
End Sub
