VERSION 5.00
Begin VB.UserControl ucNavBar 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ScaleHeight     =   5655
   ScaleWidth      =   5550
   Begin VB.PictureBox lpItem 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
End
Attribute VB_Name = "ucNavBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const WM_SETREDRAW As Long = &HB&

Public Enum eItemState
    IS_NORMAL = 0
    IS_Hot = 1
    IS_Press = 2
    IS_Selected = 3
End Enum

Public Event OnItemPrePaint(ByVal Item As Integer, ByVal ItemState As eItemState)
Public Event OnItemPostPaint(ByVal Item As Integer, ByVal Hdc As Long, ByVal ItemState As eItemState, ByVal Width As Long, Height As Long)
Public Event Click(ByVal Item As Integer)
Public Event MouseDown(ByVal Item As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(ByVal Item As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(ByVal Item As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ItemMouseEnter(ByVal Item As Integer)
Public Event ItemMouseLeave(ByVal Item As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

Private Type tNodeList
    NodeName As String
    ParentName As String
    Caption As String
    lpIndex As Integer
    ParentIndex As Integer
    ChildCount As Integer
    Ident As Integer
    Visible As Boolean
    Expanded As Boolean
    Contractable As Boolean
End Type

Private tNode() As tNodeList
Private NodeCount As Long
Private mTop As Long
Private m_CaptionPaddinX As Integer
Private m_PicIconPaddinX As Integer
Private m_SelectedIndex As Integer
Private m_MarginLeft As Integer
Private m_IconSpace As Integer
Private m_IconSize As Integer
Private m_ArrowNormalColor As OLE_COLOR
Private m_ArrowHotColor As OLE_COLOR
Private m_ArrowPushedColor As OLE_COLOR
Private m_ArrowAlignLeft As Boolean
Private m_ArrowSize As Integer
Private m_ArrowStyleSolid As Boolean
Private m_ItemHeight As Integer
Private eIState As eItemState

Public Property Get NodeName(Index As Integer) As String
    NodeName = tNode(Index - 1).NodeName
End Property

Public Property Get NodeParentName(Index As Integer) As String
    NodeParentName = tNode(Index - 1).ParentName
End Property

Public Property Get IsNodeExpanded(Index As Integer) As Boolean
    IsNodeExpanded = tNode(Index - 1).Expanded
End Property

Public Property Let ExpandedNode(Index As Integer, Value As Boolean)
    tNode(Index - 1).Expanded = Value
    Refresh
End Property

Public Property Get ChildCount(Index As Integer) As Integer
    ChildCount = tNode(Index - 1).ChildCount
End Property

Public Property Get NodeIdent(Index As Integer) As Integer
    NodeIdent = tNode(Index - 1).Ident
End Property

Public Property Get NodeItemIndex(Index As Integer) As Integer
    NodeItemIndex = tNode(Index - 1).lpIndex
End Property

Public Property Get NodeCounts() As Integer
    NodeCounts = NodeCount
End Property

Public Property Get IsNodeContractable(Index As Integer) As Boolean
    IsNodeContractable = tNode(Index - 1).Contractable
End Property

Public Property Get LabelPlusItem(ByVal Index As Integer) As LabelPlus
    Set LabelPlusItem = lpItem(Index)
End Property

Public Function NodeAdd(sName As String, Caption As String, Optional ParentNode As String, Optional Contractable As Boolean = True, Optional Expanded As Boolean, Optional IconCharCode As Long, Optional PicData As Variant)
    Dim i As Long
    Dim Ident As Integer
    ReDim Preserve tNode(NodeCount)

    Load lpItem(lpItem.Count)
    
    With tNode(NodeCount)
        .NodeName = sName
        .Caption = Caption
        .Contractable = Contractable
        .Expanded = Expanded
        If Not Contractable Then .Expanded = True
        .lpIndex = lpItem.Count - 1
       
        If Len(ParentNode) Then
            For i = 0 To NodeCount - 1
                If tNode(i).NodeName = ParentNode Then
                    .ParentIndex = i
                    .Ident = tNode(i).Ident + 1
                    tNode(i).ChildCount = tNode(i).ChildCount + 1
                    Exit For
                End If
            Next
            .ParentName = ParentNode
        Else
            .ParentIndex = -1
            .Ident = 1
        End If
    End With
    
    With lpItem(lpItem.Count - 1)
        .CaptionPaddingX = m_MarginLeft + tNode(NodeCount).Ident * m_CaptionPaddinX + m_IconSize + m_IconSpace
        .IconPaddingX = .CaptionPaddingX - m_IconSize - m_IconSpace
        .PicturePaddingX = .IconPaddingX

        .Caption = Caption
        .IconCharCode = IconCharCode
        
        If VarType(PicData) = (vbArray Or vbByte) Then
            Dim bDATA() As Byte
            bDATA = PicData
            Call .PictureFromStream(bDATA)
        End If
    End With
    
    NodeCount = NodeCount + 1
     
End Function


Private Function Recursive(Index As Integer)
    Dim i As Integer
    With lpItem(tNode(Index).lpIndex)
        If tNode(Index).ParentIndex = -1 Then
            tNode(Index).Visible = True
            .Visible = True
            .Top = mTop
            .Tag = mTop
            mTop = mTop + .Height
        Else
            If tNode(tNode(Index).ParentIndex).Expanded And tNode(tNode(Index).ParentIndex).Visible = True Then
                tNode(Index).Visible = True
                .Visible = True
                .Top = mTop
                .Tag = mTop
                mTop = mTop + .Height
            Else
                .Visible = False
                tNode(Index).Visible = False
            End If
        End If
    End With
    
    For i = 0 To NodeCount - 1
        If tNode(i).ParentName = tNode(Index).NodeName Then
            Recursive i
        End If
    Next
End Function



Public Sub Refresh()
    Dim mValue As Long
    
    Dim i As Integer, j As Long
    
    Call SendMessage(UserControl.hwnd, WM_SETREDRAW, 0&, 0&)
    
    

    ucScrollbar1.Value = 0
    mTop = 0
    For i = 0 To NodeCount - 1
        If tNode(i).ParentIndex = -1 Then
            
            Recursive i
        End If
    Next

    CheckScroll
    
    If mTop > UserControl.ScaleHeight Then
        ucScrollbar1.Value = mValue
    End If
    
  Call SendMessage(UserControl.hwnd, WM_SETREDRAW, 1&, 0&)
  RedrawWindow UserControl.hwnd, ByVal &H0, 0, 1
  ucScrollbar1.Refresh
End Sub

Private Sub CheckScroll()
    Dim i As Long
    If mTop > UserControl.ScaleHeight Then
        ucScrollbar1.SmallChange = lpItem(0).Height
        ucScrollbar1.Max = mTop - UserControl.ScaleHeight
        ucScrollbar1.LargeChange = mTop - UserControl.ScaleHeight
        For i = 0 To lpItem.Count - 1
            lpItem(i).Width = UserControl.ScaleWidth - ucScrollbar1.Width
        Next
        ucScrollbar1.Visible = True
    Else
        ucScrollbar1.Visible = False
        For i = 0 To lpItem.Count - 1
            lpItem(i).Top = Val(lpItem(i).Tag)
            lpItem(i).Width = UserControl.ScaleWidth
        Next
    End If
End Sub



Private Sub ucNavBar_Click(Index As Integer)
    Dim i As Long
    Dim LstChild As Integer
    
    If tNode(Index - 1).Contractable Then
        tNode(Index - 1).Expanded = Not tNode(Index - 1).Expanded
    End If
       
    If tNode(Index - 1).ChildCount = 0 Then
        m_SelectedIndex = Index
    End If
    
    Refresh
    
    If ucScrollbar1.Visible = True And tNode(Index - 1).Expanded = True Then
        LstChild = GetLastChild(Index - 1)
        If LstChild > -1 Then
            Call SendMessage(UserControl.hwnd, WM_SETREDRAW, 0&, 0&)
            EnsureVisible LstChild + 1
            EnsureVisible Index
            Call SendMessage(UserControl.hwnd, WM_SETREDRAW, 1&, 0&)
            RedrawWindow UserControl.hwnd, ByVal &H0, 0, 1
        End If
    End If
    
    RaiseEvent Click(Index)

End Sub


Private Function GetLastChild(Index As Integer) As Integer
    Dim i As Integer
    If tNode(Index).ChildCount > 0 Then
        For i = NodeCount - 1 To 1 Step -1
            If tNode(i).ParentName = tNode(Index).NodeName Then
                GetLastChild = i
                Exit Function
            End If
        Next
    End If
    GetLastChild = -1
End Function

Private Sub ucNavBar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lpItem_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lpItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lpItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lpItem(Index).Refresh
    RaiseEvent MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lpItem_MouseEnter(Index As Integer)
    lpItem(Index).Refresh
    RaiseEvent ItemMouseEnter(Index)
End Sub

Private Sub lpItem_MouseLeave(Index As Integer)
    lpItem(Index).Refresh
    RaiseEvent ItemMouseLeave(Index)
End Sub

Private Sub lpItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lpItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lpItem(Index).Refresh
    RaiseEvent MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lpItem_PostPaint(Index As Integer, ByVal Hdc As Long)
    Dim H As Long
    Dim Left As Long
    Dim ArrowColor As Long
    Dim ASZ As Integer
    If tNode(Index - 1).ChildCount > 0 And tNode(Index - 1).Contractable Then
        
        With lpItem(Index)

            H = .Height \ Screen.TwipsPerPixelY \ .GetWindowsDPI
            
            If m_ArrowAlignLeft Then
                Left = m_MarginLeft + tNode(Index - 1).Ident * m_CaptionPaddinX - m_CaptionPaddinX
            Else
                Left = (.Width \ Screen.TwipsPerPixelX \ .GetWindowsDPI) - m_ArrowSize - (m_IconSpace * 2)
            End If
            
            ArrowColor = IIf(.IsMouseEnter, m_ArrowHotColor, m_ArrowNormalColor)
            ASZ = m_ArrowSize
            If m_ArrowStyleSolid Then
                If tNode(Index - 1).Expanded Then
                    .Polygon Hdc, 0, ArrowColor, 100, Left - ASZ, H \ 2 - ASZ \ 2, Left + ASZ, H \ 2 - ASZ \ 2, Left, H \ 2 + ASZ \ 2 '+ ASZ \ 2
                Else
                    .Polygon Hdc, 0, ArrowColor, 100, Left, H \ 2 - ASZ, Left + ASZ, H \ 2, Left, H \ 2 + ASZ
                End If
            Else
                If tNode(Index - 1).Expanded Then
                    .DrawLine Hdc, Left - ASZ, H \ 2 - ASZ \ 2, Left, H \ 2 + ASZ \ 2, ArrowColor, , 1
                    .DrawLine Hdc, Left, H \ 2 + ASZ \ 2, Left + ASZ, H \ 2 - ASZ \ 2, ArrowColor, , 1
                Else
                    .DrawLine Hdc, Left, H \ 2 - ASZ, Left + ASZ, H \ 2, ArrowColor, , 1
                    .DrawLine Hdc, Left + ASZ, H \ 2, Left, H \ 2 + ASZ, ArrowColor, , 1
                End If
            End If
        End With
    End If
    
    RaiseEvent OnItemPostPaint(Index, Hdc, eIState, lpItem(Index).Width, lpItem(Index).Height)

End Sub

Public Sub EnsureVisible(Index)
    If lpItem(Index).Top + lpItem(Index).Height > UserControl.ScaleHeight Then
        'Do While lpItem(Index).Top + lpItem(Index).Height > UserControl.ScaleHeight
        '    ucScrollbar1.Value = ucScrollbar1.Value + lpItem(Index).Height
        '    DoEvents
        'Loop
        ucScrollbar1.Value = ucScrollbar1.Value - UserControl.ScaleHeight + (lpItem(Index).Top + lpItem(Index).Height)
    End If
    
    If lpItem(Index).Top < 0 Then
        'Do While lpItem(Index).Top < 0
            'ucScrollbar1.Value = ucScrollbar1.Value - lpItem(Index).Height
        '    DoEvents
       ' Loop
        ucScrollbar1.Value = ucScrollbar1.Value + lpItem(Index).Top
    End If

End Sub

Private Sub lpItem_PrePaint(Index As Integer, Hdc As Long, X As Long, Y As Long)
    
    eIState = IS_NORMAL
    If lpItem(Index).IsMouseEnter Then eIState = IS_Hot
    
    If GetKeyState(vbLeftButton) < 0 And lpItem(Index).IsMouseEnter Then eIState = IS_Press
    If Index = m_SelectedIndex Then eIState = IS_Selected
    lpItem(Index).Redraw = False
    RaiseEvent OnItemPrePaint(Index, eIState)
    lpItem(Index).Redraw = True
End Sub

Private Sub UserControl_Initialize()
    m_CaptionPaddinX = 10
End Sub


Private Sub UserControl_InitProperties()
    m_ArrowSize = 4
    m_IconSpace = 8
    m_IconSize = 16
    m_ArrowAlignLeft = False
    m_MarginLeft = 0
    m_ItemHeight = 735
    m_ArrowNormalColor = vbHighlight
    m_ArrowHotColor = vbHighlight
    m_ArrowPushedColor = vbHighlight
    UserControl.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    ucScrollbar1.Move UserControl.ScaleWidth - ucScrollbar1.Width, 0, ucScrollbar1.Width, UserControl.ScaleHeight
    Refresh
End Sub

Private Sub ucScrollbar1_Change()
    Dim i As Long
    For i = 0 To lpItem.Count - 1
        lpItem(i).Top = Val(lpItem(i).Tag) - ucScrollbar1.Value
    Next
    DoEvents
End Sub

Private Sub ucScrollbar1_Scroll()
    ucScrollbar1_Change
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor, Ambient.BackColor)
        Call .WriteProperty("Font", lpItem(0).Font, UserControl.Ambient.Font)
        Call .WriteProperty("PictureSize", lpItem(0).PictureSetWidth, 16)
        Call .WriteProperty("IconFont", lpItem(0).IconFont, UserControl.Ambient.Font)
        Call .WriteProperty("IconSize", m_IconSize, 16)
        Call .WriteProperty("MarginLeft", m_MarginLeft, 0)
        Call .WriteProperty("IconSpace", m_IconSpace, 8)
        Call .WriteProperty("ArrowNormalColor", m_ArrowNormalColor, vbHighlight)
        Call .WriteProperty("ArrowHotColor", m_ArrowHotColor, vbHighlight)
        Call .WriteProperty("ArrowPushedColor", m_ArrowPushedColor, vbHighlight)
        Call .WriteProperty("ArrowSize", m_ArrowSize, 4)
        Call .WriteProperty("ArrowStyleSolid", m_ArrowStyleSolid, False)
        Call .WriteProperty("ItemHeight", m_ItemHeight, 735)
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        Set lpItem(0).Font = .ReadProperty("Font", UserControl.Ambient.Font)
        lpItem(0).PictureSetWidth = .ReadProperty("PictureSize", 16)
        Set lpItem(0).IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_MarginLeft = .ReadProperty("MarginLeft", 0)
        m_IconSpace = .ReadProperty("IconSpace", 8)
        m_ArrowNormalColor = .ReadProperty("ArrowNormalColor", vbHighlight)
        m_ArrowHotColor = .ReadProperty("ArrowHotColor", vbHighlight)
        m_ArrowPushedColor = .ReadProperty("ArrowPushedColor", vbHighlight)
        m_ArrowAlignLeft = .ReadProperty("ArrowAlignLeft", False)
        m_ArrowSize = .ReadProperty("ArrowSize", 4)
        m_ArrowStyleSolid = .ReadProperty("ArrowStyleSolid", False)
        m_ItemHeight = .ReadProperty("ItemHeight", 735)
    End With
    
    lpItem(0).Height = m_ItemHeight
    lpItem(0).PictureSetWidth = m_IconSize
    
    ucScrollbar1.BackColor = UserControl.BackColor
    ucScrollbar1.TrackColor = UserControl.BackColor
    ucScrollbar1.ThemeColor = UserControl.BackColor
End Sub


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    ucScrollbar1.BackColor = UserControl.BackColor
    ucScrollbar1.TrackColor = UserControl.BackColor
    ucScrollbar1.ThemeColor = UserControl.BackColor
    PropertyChanged "BackColor"
    'Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = lpItem(0).Font
End Property

Public Property Set Font(New_Font As StdFont)
    Set lpItem(0).Font = New_Font
    PropertyChanged "Font"
    'Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = lpItem(0).IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
    Set lpItem(0).IconFont = New_Font
    'm_IconSize = lpItem(0).Font.Size 'Aproximate
    PropertyChanged "IconFont"
    'Refresh
End Property

Public Property Get IconSpace() As Integer
    IconSpace = m_IconSpace
End Property

Public Property Let IconSpace(ByVal New_Value As Integer)
    m_IconSpace = New_Value
    PropertyChanged "IconSpace"
    'Refresh
End Property

Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Value As Integer)
    m_IconSize = New_Value
    PropertyChanged "IconSize"
    'Refresh
End Property

Public Property Get MarginLeft() As Integer
    MarginLeft = m_MarginLeft
End Property

Public Property Let MarginLeft(ByVal New_Value As Integer)
    m_MarginLeft = New_Value
    PropertyChanged "MarginLeft"
    'Refresh
End Property

Public Property Get ArrowNormalColor() As OLE_COLOR
    ArrowNormalColor = m_ArrowNormalColor
End Property

Public Property Let ArrowNormalColor(ByVal New_Value As OLE_COLOR)
    m_ArrowNormalColor = New_Value
    PropertyChanged "ArrowNormalColor"
    'Refresh
End Property

Public Property Get ArrowHotColor() As OLE_COLOR
    ArrowHotColor = m_ArrowHotColor
End Property

Public Property Let ArrowHotColor(ByVal New_Value As OLE_COLOR)
    m_ArrowHotColor = New_Value
    PropertyChanged "ArrowHotColor"
    'Refresh
End Property

Public Property Get ArrowPushedColor() As OLE_COLOR
    ArrowPushedColor = m_ArrowPushedColor
End Property

Public Property Let ArrowPushedColor(ByVal New_Value As OLE_COLOR)
    m_ArrowPushedColor = New_Value
    PropertyChanged "ArrowPushedColor"
    'Refresh
End Property

Public Property Get ArrowAlignLeft() As Boolean
    ArrowAlignLeft = m_ArrowAlignLeft
End Property

Public Property Let ArrowAlignLeft(ByVal New_Value As Boolean)
    m_ArrowAlignLeft = New_Value
    PropertyChanged "ArrowAlignLeft"
    'Refresh
End Property

Public Property Get ArrowSize() As Integer
    ArrowSize = m_ArrowSize
End Property

Public Property Let ArrowSize(ByVal New_Value As Integer)
    m_ArrowSize = New_Value
    PropertyChanged "ArrowSize"
    'Refresh
End Property

Public Property Get ArrowStyleSolid() As Boolean
    ArrowStyleSolid = m_ArrowStyleSolid
End Property

Public Property Let ArrowStyleSolid(ByVal New_Value As Boolean)
    m_ArrowStyleSolid = New_Value
    PropertyChanged "ArrowStyleSolid"
    'Refresh
End Property

Public Property Get ItemHeight() As Integer
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_Value As Integer)
    m_ItemHeight = New_Value
    lpItem(0).Height = New_Value
    PropertyChanged "ItemHeight"
    'Refresh
End Property



