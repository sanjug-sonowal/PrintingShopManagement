VERSION 5.00
Begin VB.UserControl ucNeumorphism 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   1320
      Top             =   1800
   End
End
Attribute VB_Name = "ucNeumorphism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: LabelPlus
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'LastUpdate: 14/02/2020
'Version: 1.5.3
'Based on: FirenzeLabel Project :http://www.vbforums.com/showthread.php?845221-VB6-FIRENZE-LABEL-label-control-with-so-many-functions
           'Martin Vartiak, powered by Cairo Graphics and vbRichClient-Framework.
'Special thanks to: All members of the VB6 Latin group (www.leandroacierto.com/foro), vbforum.com and activevb.net
'-----------------------------------------------

Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GdipCloneBrush Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mCloneBrush As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipTranslateClipI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mDx As Long, ByVal mDy As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal Path As Long, ByVal x As Single, ByVal y As Single, ByVal graphics As Long, result As Long) As Long
Private Declare Function GdipClonePath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mClonePath As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Const UnitPixel                 As Long = &H2&
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CombineModeExclude        As Long = &H4
Private Const WrapModeTileFlipXY        As Long = &H3
Private Const IDC_HAND                  As Long = 32649
Private Const LOGPIXELSX                As Long = 88
'Private Const LOGPIXELSY                As Long = 90
Private Const WM_MOUSEMOVE              As Long = &H200

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type


Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Public Enum eLightDirection
    TopLeft = 0
    TopRight = 1
    BottomRight = 2
    BottomLeft = 3
End Enum

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseOver()
Public Event MouseOut()
Public Event PrePaint(hdc As Long, x As Long, y As Long)
Public Event PostPaint(ByVal hdc As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'Public Event PictureDownloadProgress(BytesMax As Long, BytesLeidos As Long)
'Public Event PictureDownloadComplete()
'Public Event PictureDownloadError()


Public m_BrushBtnNormal As Long
Public m_BrushBtnPressed As Long

Dim GdipToken As Long
Dim m_Distance As Long
Dim m_Radius As Long
Dim m_Intencity As Long
Dim m_Blur As Long
Dim m_BackColor As OLE_COLOR
Dim m_LightDirection As eLightDirection
Dim m_StatePressed As Boolean
Dim m_Gradient As Boolean
Dim m_GradientFlip As Boolean
Dim m_ShadowColor As OLE_COLOR
Dim m_LightColor As OLE_COLOR
Dim m_Invert As Boolean
Dim m_Redraw As Boolean
Dim m_MousePointerHands As Boolean
Dim m_MouseToParent As Boolean
Dim hCur As Long
Dim c_lhWnd As Long
Dim nScale As Single
Dim bIntercept As Boolean
Dim m_Enter As Boolean
Dim m_Over As Boolean
Dim m_PT As POINTAPI
Dim m_Left As Long
Dim m_Top As Long
Dim m_ToggleButtonStyle As Boolean
Dim m_ButtonStyle As Boolean
Dim IsPressed As Boolean
Dim mPath As Long


Public Function CloneTo(ByVal SRC_ucNeumorphism As ucNeumorphism, Optional TwoState As Boolean, Optional DontClone As Boolean)
    Dim hCopyBrush As Long
    With SRC_ucNeumorphism
       
        Draw UserControl.hdc, 0, 0
        
        If TwoState Then
            m_StatePressed = Not m_StatePressed
            Draw UserControl.hdc, 0, 0
            m_StatePressed = Not m_StatePressed
        End If

        If m_BrushBtnNormal <> 0 Then
            If .m_BrushBtnNormal <> 0 Then GdipDeleteBrush .m_BrushBtnNormal
            If DontClone Then
                .m_BrushBtnNormal = m_BrushBtnNormal 'this needs to be reviewed, just as proof
            Else
                Call GdipCloneBrush(m_BrushBtnNormal, hCopyBrush)
                .m_BrushBtnNormal = hCopyBrush
            End If
        End If
        
        If m_BrushBtnPressed <> 0 Then
            If .m_BrushBtnPressed <> 0 Then GdipDeleteBrush .m_BrushBtnPressed
            If DontClone Then
                .m_BrushBtnPressed = m_BrushBtnPressed 'this needs to be reviewed, just as proof
            Else
                GdipCloneBrush m_BrushBtnPressed, hCopyBrush
                .m_BrushBtnPressed = hCopyBrush
            End If
        End If

    End With
End Function

Public Property Get Invert() As Boolean
    Invert = m_Invert
End Property

Public Property Let Invert(ByVal new_value As Boolean)
    m_Invert = new_value
    PropertyChanged "Invert"
End Property

Public Property Get ButtonStyle() As Boolean
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal new_value As Boolean)
    m_ButtonStyle = new_value
    PropertyChanged "ButtonStyle"
End Property

Public Property Get ToggleButtonStyle() As Boolean
    ToggleButtonStyle = m_ToggleButtonStyle
End Property

Public Property Let ToggleButtonStyle(ByVal new_value As Boolean)
    m_ToggleButtonStyle = new_value
    PropertyChanged "ToggleButtonStyle"
End Property

Public Property Get Distance() As Long
    Distance = m_Distance
End Property

Public Property Let Distance(ByVal new_value As Long)
    m_Distance = new_value
    CleanUp
    PropertyChanged "Distance"
    Refresh
End Property

Public Property Get Radius() As Long
    Radius = m_Radius
End Property

Public Property Let Radius(ByVal new_value As Long)
    m_Radius = new_value
    CleanUp
    PropertyChanged "Radius"
    Refresh
End Property

Public Property Get Intencity() As Long
    Intencity = m_Intencity
End Property

Public Property Let Intencity(ByVal new_value As Long)
    m_Intencity = new_value
    CleanUp
    PropertyChanged "Intencity"
    Refresh
End Property

Public Property Get Blur() As Long
    Blur = m_Blur
End Property

Public Property Let Blur(ByVal new_value As Long)
    m_Blur = new_value
    CleanUp
    PropertyChanged "Blur"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal new_value As OLE_COLOR)
    m_BackColor = new_value
    CleanUp
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal new_value As OLE_COLOR)
    m_ShadowColor = new_value
    CleanUp
    PropertyChanged "ShadowColor"
    Refresh
End Property

Public Property Get LightColor() As OLE_COLOR
    LightColor = m_LightColor
End Property

Public Property Let LightColor(ByVal new_value As OLE_COLOR)
    m_LightColor = new_value
    CleanUp
    PropertyChanged "LightColor"
    Refresh
End Property
Public Property Get LightDirection() As eLightDirection
    LightDirection = m_LightDirection
End Property

Public Property Let LightDirection(ByVal new_value As eLightDirection)
    m_LightDirection = new_value
    CleanUp
    PropertyChanged "LightDirection"
    Refresh
End Property

Public Property Get StatePressed() As Boolean
    StatePressed = m_StatePressed
End Property

Public Property Let StatePressed(ByVal new_value As Boolean)
    m_StatePressed = new_value
    PropertyChanged "StatePressed"
    Refresh
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal new_value As Boolean)
    m_Gradient = new_value
    CleanUp
    PropertyChanged "Gradient"
    Refresh
End Property

Public Property Get GradientFlip() As Boolean
    GradientFlip = m_GradientFlip
End Property

Public Property Let GradientFlip(ByVal new_value As Boolean)
    m_GradientFlip = new_value
    CleanUp
    PropertyChanged "GradienFlip"
    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

Public Property Get MouseToParent() As Boolean
    MouseToParent = m_MouseToParent
End Property

Public Property Let MouseToParent(ByVal new_value As Boolean)
    m_MouseToParent = new_value
    PropertyChanged "MouseToParent"
End Property

Public Property Let OLEDropMode(ByVal new_value As OLEDropConstants)
    UserControl.OLEDropMode = new_value
End Property
Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = UserControl.OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal new_value As Boolean)
    m_Redraw = new_value
End Property


Public Sub CleanUp()
    If m_BrushBtnNormal Then GdipDeleteBrush m_BrushBtnNormal: m_BrushBtnNormal = 0&
    If m_BrushBtnPressed Then GdipDeleteBrush m_BrushBtnPressed: m_BrushBtnPressed = 0&
    If mPath <> 0 Then GdipDeletePath mPath: mPath = 0&
End Sub

'*1
Public Sub Draw(ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional CustomPath As Long)
    Dim hGraphics As Long, hGraphics2 As Long
    Dim hPath As Long, hBrush As Long, hPen As Long
    Dim hImage As Long
    Dim IsDark As Boolean
    Dim DB2 As Long, B2 As Long
    Dim GradientAngle As Long
    Dim x As Long, y As Long
    Dim LC As Long
    Dim RECT As RECTL
    Dim Color1 As Long, Color2 As Long
    Dim D1 As Long, B As Long
    
    If Width = 0 Then Width = UserControl.ScaleWidth - ((m_Distance + m_Blur * 2) * 2) * nScale
    If Height = 0 Then Height = UserControl.ScaleHeight - ((m_Distance + m_Blur * 2) * 2) * nScale
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    D1 = m_Distance * nScale
    B = m_Blur * nScale
    B2 = B * 2 ' + 1
    DB2 = (B2 + D1) * 2
        
    If m_BrushBtnNormal And (m_StatePressed = False) Then
        GdipTranslateWorldTransform hGraphics, Left - D1 - B2, Top - D1 - B2, 0&
        GdipFillRectangleI hGraphics, m_BrushBtnNormal, 0, 0, Width + DB2, Height + DB2
        GdipDeleteGraphics hGraphics
        Exit Sub
    End If
    
    If m_BrushBtnPressed And (m_StatePressed = True) Then
        GdipTranslateWorldTransform hGraphics, Left, Top, 0&
        GdipFillPath hGraphics, m_BrushBtnPressed, mPath
        GdipDeleteGraphics hGraphics
        Exit Sub
    End If
    
    If mPath <> 0 Then GdipDeletePath mPath: mPath = 0
    
    If CustomPath Then
        GdipClonePath CustomPath, hPath
    Else
        hPath = CreateRoundPath(0, 0, Width, Height, m_Radius * nScale)
    End If
    
    If hPath Then
        mPath = hPath
    Else
        GdipDeleteGraphics hGraphics
        Exit Sub
    End If
    
    IsDark = IsDarkColor(m_BackColor)
    LC = LuminanceColor(m_BackColor)
    
    If m_StatePressed = False Then
        GdipCreateBitmapFromScan0 Width + DB2, Height + DB2, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
        GdipGetImageGraphicsContext hImage, hGraphics2
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
    
        x = B2: y = B2
        Select Case m_LightDirection
            Case TopRight: x = x + D1 * 2
            Case BottomRight: x = x + D1 * 2: y = x
            Case BottomLeft: y = y + D1 * 2
        End Select
        
        GdipTranslateWorldTransform hGraphics2, x, y, 0&
        GdipCreateSolidFill RGBtoARGB(m_LightColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush

        x = B2: y = B2
        Select Case m_LightDirection
            Case TopLeft: x = x + D1 * 2: y = x
            Case TopRight: y = y + D1 * 2
            Case BottomLeft: x = x + D1 * 2
        End Select
        
        GdipResetWorldTransform hGraphics2
        GdipTranslateWorldTransform hGraphics2, x, y, 0&
        GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
    
        BlurImage hImage, B
    
        GdipResetWorldTransform hGraphics2
        GdipTranslateWorldTransform hGraphics2, D1 + B2, D1 + B2, 0&
        If m_Gradient Then
            RECT.Width = Width
            RECT.Height = Height
            GradientAngle = 45 + 90 * m_LightDirection - 180 * m_GradientFlip
            Color1 = RGBtoARGB(ShiftColor(m_ShadowColor, m_BackColor, m_Intencity), 100)
            Color2 = RGBtoARGB(ShiftColor(m_LightColor, m_BackColor, m_Intencity), 100)
            GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, GradientAngle, 0, WrapModeTileFlipXY, hBrush
        Else
            GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
            GdipFillPath hGraphics2, hBrush, hPath
        End If
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
        
        GdipCreateTexture hImage, &H0, hBrush
        GdipTranslateWorldTransform hGraphics, Left - D1 - B2, Top - D1 - B2, 0&
        GdipFillRectangleI hGraphics, hBrush, 0, 0, Width + DB2, Height + DB2
        If m_BrushBtnNormal Then GdipDeleteBrush m_BrushBtnNormal 'cleanup the last
        m_BrushBtnNormal = hBrush
    Else
'*2
        Dim hImage2 As Long, hGraphics3 As Long
        GdipCreateBitmapFromScan0 Width + B * 2, Height + B * 2, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
        GdipGetImageGraphicsContext hImage, hGraphics2
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
        GdipCreateBitmapFromScan0 Width + B * 2, Height + B * 2, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics3
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
        
        If m_Gradient Then
            RECT.Width = Width
            RECT.Height = Height
            GradientAngle = 45 + 90 * m_LightDirection - 180 * m_GradientFlip
            Color1 = RGBtoARGB(ShiftColor(m_ShadowColor, m_BackColor, m_Intencity), 100)
            Color2 = RGBtoARGB(ShiftColor(m_LightColor, m_BackColor, m_Intencity), 100)
            GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, GradientAngle, 0, WrapModeTileFlipXY, hBrush
        Else
            GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
        End If
        GdipFillPath hGraphics3, hBrush, hPath
        GdipDeleteBrush hBrush
            
        x = -D1: y = -D1
        Select Case m_LightDirection
            Case TopRight: x = x + D1 * 2
            Case BottomRight: x = x + D1 * 2: y = x
            Case BottomLeft: y = y + D1 * 2
        End Select
       
        GdipSetClipPath hGraphics2, hPath, CombineModeExclude
        GdipTranslateClipI hGraphics2, x, y
        GdipCreateSolidFill RGBtoARGB(m_LightColor, m_Intencity), hBrush 'IIf(IsDark, m_Intencity, 50)
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
    
        GdipTranslateClipI hGraphics2, (x * -1) * 2, (y * -1) * 2
        GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
    
        If IsDark = False Then
            GdipResetClip hGraphics2
            GdipCreatePen1 RGBtoARGB(m_ShadowColor, m_Intencity / 2), B / 4, UnitPixel, hPen
            GdipDrawPath hGraphics2, hPen, hPath
            GdipDeletePen hPen
        End If
    
        BlurImage hImage, B
        
        GdipDrawImageRectI hGraphics3, hImage, 0, 0, Width + B * 2, Height + B * 2
        GdipCreateTexture hImage2, &H0, hBrush
        GdipTranslateWorldTransform hGraphics, Left, Top, 0&
        GdipFillPath hGraphics, hBrush, hPath
        If m_BrushBtnPressed Then GdipDeleteBrush m_BrushBtnPressed 'cleanup the last
        m_BrushBtnPressed = hBrush
    End If
    
    'If CustomPath = 0 Then GdipDeletePath hPath
    GdipDisposeImage hImage2
    GdipDeleteGraphics hGraphics3
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics2
    GdipDeleteGraphics hGraphics

End Sub

Private Function CreateRoundPath(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Radius As Single) As Long
    Dim hPath As Long
    If GdipCreatePath(&H0, hPath) = 0& Then
    
        If Radius > Width / 2 Then Radius = Width / 2
        If Radius > Height / 2 Then Radius = Height / 2
    
        If Radius = 0 Then
            GdipAddPathRectangleI hPath, Left, Top, Width, Height
        Else
            Radius = Radius * 2
            GdipAddPathArcI hPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI hPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure hPath
        End If
        CreateRoundPath = hPath
    End If
    
End Function


Private Function BlurImage(ByVal hImage As Long, blurDepth As Long, _
                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Boolean
                                        
    Dim RECT As RECTL
    Dim bmpData1 As BitmapData
    Dim srcBytes() As Byte
    Dim kDiv As Long
    Dim MaxX As Long, MaxY As Long, Width4 As Long
    Dim x As Long, y As Long
    Dim X0 As Long, X1 As Long, X2 As Long
    Dim Y0 As Long, Y1 As Long, Y2 As Long
    Dim A As Long, R As Long, G As Long, B As Long
    Dim dX0 As Long, dX2 As Long, dY0 As Long, dY2 As Long
    Dim mOut() As Byte, n As Long

    If blurDepth <= 0& Then Exit Function
    If hImage = 0& Then Exit Function
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    With RECT
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With

    ReDim srcBytes(RECT.Width * RECT.Height * 4 - 1&)

    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&))
        .stride = 4& * RECT.Width
    End With
   
    If GdipBitmapLockBits(hImage, RECT, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData1) = 0& Then

        MaxX = Width - 1
        MaxY = Height - 1
        Width4 = Width * 4
        ReDim mOut(Width4 * Height - 1)
        kDiv = blurDepth * 2 + 1
    
        For n = 0 To 1
          For y = 0 To MaxY
            B = 0
            G = 0
            R = 0
            A = 0
            X0 = y * Width4
            X1 = X0
            For x = 2 To blurDepth
              X0 = X0 + 4
              B = B + srcBytes(X0 + 0)
              G = G + srcBytes(X0 + 1)
              R = R + srcBytes(X0 + 2)
              A = A + srcBytes(X0 + 3)
            Next x
            X0 = X1 + blurDepth * 4
            X2 = X0
            B = B + B + srcBytes(X1 + 0) + srcBytes(X0 + 0)
            G = G + G + srcBytes(X1 + 1) + srcBytes(X0 + 1)
            R = R + R + srcBytes(X1 + 2) + srcBytes(X0 + 2)
            A = A + A + srcBytes(X1 + 3) + srcBytes(X0 + 3)
            dX0 = -4
            dX2 = 4
            For x = 0 To MaxX
              B = B + srcBytes(X2 + 0)
              G = G + srcBytes(X2 + 1)
              R = R + srcBytes(X2 + 2)
              A = A + srcBytes(X2 + 3)
              mOut(X1 + 0) = B \ kDiv
              mOut(X1 + 1) = G \ kDiv
              mOut(X1 + 2) = R \ kDiv
              mOut(X1 + 3) = A \ kDiv
              B = B - srcBytes(X0 + 0)
              G = G - srcBytes(X0 + 1)
              R = R - srcBytes(X0 + 2)
              A = A - srcBytes(X0 + 3)
              If x = blurDepth Then dX0 = 4
              X0 = X0 + dX0
              X1 = X1 + 4
              If x = MaxX - blurDepth Then dX2 = -4
              X2 = X2 + dX2
            Next x
          Next y
          
          For x = 0 To MaxX
            B = 0
            G = 0
            R = 0
            A = 0
            Y0 = x * 4
            Y1 = Y0
            For y = 2 To blurDepth
              Y0 = Y0 + Width4
              B = B + mOut(Y0 + 0)
              G = G + mOut(Y0 + 1)
              R = R + mOut(Y0 + 2)
              A = A + mOut(Y0 + 3)
            Next y
            Y0 = Y1 + blurDepth * Width4
            Y2 = Y0
            B = B + B + mOut(Y1 + 0) + mOut(Y0 + 0)
            G = G + G + mOut(Y1 + 1) + mOut(Y0 + 1)
            R = R + R + mOut(Y1 + 2) + mOut(Y0 + 2)
            A = A + A + mOut(Y1 + 3) + mOut(Y0 + 3)
            dY0 = -Width4
            dY2 = Width4
            For y = 0 To MaxY
              B = B + mOut(Y2 + 0)
              G = G + mOut(Y2 + 1)
              R = R + mOut(Y2 + 2)
              A = A + mOut(Y2 + 3)
              srcBytes(Y1 + 0) = B \ kDiv
              srcBytes(Y1 + 1) = G \ kDiv
              srcBytes(Y1 + 2) = R \ kDiv
              srcBytes(Y1 + 3) = A \ kDiv
              B = B - mOut(Y0 + 0)
              G = G - mOut(Y0 + 1)
              R = R - mOut(Y0 + 2)
              A = A - mOut(Y0 + 3)
              If y = blurDepth Then dY0 = Width4
              Y0 = Y0 + dY0
              Y1 = Y1 + Width4
              If y = MaxY - blurDepth Then dY2 = -Width4
              Y2 = Y2 + dY2
            Next y
          Next x
    
        Next n
      
        BlurImage = GdipBitmapUnlockBits(hImage, bmpData1) = 0&
    End If
End Function

Public Function LuminanceColor(ByVal Color As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
    CopyMemory BGRA(0), Color, 4&
    LuminanceColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) * 100 / 638
    'LuminanceColor = ((CLng(BGRA(0)) + (CLng(BGRA(1))) + CLng(BGRA(2))) / 2) * 100 / 382
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

'Funcion para combinar dos colores
Public Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function


'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub
  
'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub



Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function


Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    On Error Resume Next

    If UserControl.Enabled Then
        If Not MouseToParent Then
            HitResult = vbHitResultHit
        ElseIf Not Ambient.UserMode Then
            HitResult = vbHitResultHit
        End If
        If Ambient.UserMode Then
            Dim PT As POINTAPI
            Dim lHwnd As Long
            GetCursorPos PT
            lHwnd = WindowFromPoint(PT.x, PT.y)
            
            If m_Enter = False Then

                ScreenToClient c_lhWnd, PT
                m_PT.x = PT.x - x
                m_PT.y = PT.y - y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
 
                m_Enter = True
                tmrMOUSEOVER.Interval = 1
                 RaiseEvent MouseEnter
            End If
        
            bIntercept = True
            
            If lHwnd = c_lhWnd Then
                If m_Over = False Then
                    m_Over = True
                    RaiseEvent MouseOver
                End If
            Else
                If m_Over = True Then
                    m_Over = False
                    RaiseEvent MouseOut
                End If
            End If
        End If
    ElseIf Not Ambient.UserMode Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    nScale = GetWindowsDPI
    m_Redraw = True
    InitGDI
    'm_ToggleButtonStyle = True
    
End Sub

Private Sub UserControl_InitProperties()
   ' hFontCollection = ReadValue(&HFC)
   
    c_lhWnd = UserControl.ContainerHwnd
       
    m_Distance = 10
    m_Radius = 10
    m_Intencity = 40
    m_Blur = 15
    m_ShadowColor = vbBlack
    m_LightColor = vbWhite
    m_BackColor = Ambient.BackColor
    m_ButtonStyle = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_Paint()
    Dim lHdc As Long
    Dim x As Long, y As Long
    
    lHdc = UserControl.hdc
    x = (m_Distance + m_Blur * 2) * nScale
    y = (m_Distance + m_Blur * 2) * nScale

    RaiseEvent PrePaint(lHdc, x, y)
    'Call Draw(lHdc, 0, X, Y)
    Draw lHdc, x, y
    RaiseEvent PostPaint(UserControl.hdc)
End Sub

Public Sub Refresh()
    If m_Redraw Then
        UserControl.Refresh
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    c_lhWnd = UserControl.ContainerHwnd

    With PropBag
        m_Invert = .ReadProperty("Invert", False)
        m_Distance = .ReadProperty("Distance", 10)
        m_Radius = .ReadProperty("Radius", 10)
        m_Intencity = .ReadProperty("Intencity", 40)
        m_Blur = .ReadProperty("Blur", 15)
        m_LightDirection = .ReadProperty("LightDirection", TopLeft)
        m_StatePressed = .ReadProperty("StatePressed", False)
        m_Gradient = .ReadProperty("Gradient", False)
        m_GradientFlip = .ReadProperty("GradientFlip", False)
        m_ShadowColor = .ReadProperty("ShadowColor", vbBlack)
        m_LightColor = .ReadProperty("LightColor", vbWhite)
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        m_MouseToParent = .ReadProperty("MouseToParent", False)
        UserControl.OLEDropMode = .ReadProperty("OLEDropMode", 0&)
        m_ButtonStyle = .ReadProperty("ButtonStyle", True)
        m_ToggleButtonStyle = .ReadProperty("ToggleButtonStyle", False)
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If
    
        If m_Invert Then IsPressed = Not IsPressed
    End With
  
End Sub

Private Sub UserControl_Resize()
    CleanUp
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
         
        Call .WriteProperty("Invert", m_Invert, False)
        Call .WriteProperty("Distance", m_Distance, 10)
        Call .WriteProperty("Radius", m_Radius, 10)
        Call .WriteProperty("Intencity", m_Intencity, 40)
        Call .WriteProperty("Blur", m_Blur, 15)
        Call .WriteProperty("LightDirection", m_LightDirection, TopLeft)
        Call .WriteProperty("StatePressed", m_StatePressed, False)
        Call .WriteProperty("Gradient", m_Gradient, False)
        Call .WriteProperty("GradientFlip", m_GradientFlip, False)
        Call .WriteProperty("ShadowColor", m_ShadowColor, vbBlack)
        Call .WriteProperty("LightColor", m_LightColor, vbWhite)
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbArrow)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
        Call .WriteProperty("MousePointerHands", m_MousePointerHands, False)
        Call .WriteProperty("MouseToParent", m_MouseToParent, False)
        Call .WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0&)
        Call .WriteProperty("ButtonStyle", m_ButtonStyle, True)
        Call .WriteProperty("ToggleButtonStyle", m_ToggleButtonStyle, False)
    End With

End Sub

Private Sub UserControl_Terminate()

    If hCur Then DestroyCursor hCur
    CleanUp
    TerminateGDI
End Sub



Public Sub tmrMOUSEOVER_Timer()
    Dim PT As POINTAPI
    Dim Left As Long, Top As Long
    Dim RECT As RECT
  
    GetCursorPos PT
    ScreenToClient c_lhWnd, PT
    
    Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
    Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)

    With RECT
        .Left = m_PT.x - (m_Left - Left)
        .Top = m_PT.y - (m_Top - Top)
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
    bIntercept = False
    SendMessage c_lhWnd, WM_MOUSEMOVE, 0&, ByVal PT.x Or PT.y * &H10000
    
    If bIntercept = False Then
        If m_Over = True Then
            m_Over = False
            RaiseEvent MouseOut
        End If
    End If
    
    
    If PtInRect(RECT, PT.x, PT.y) = 0 Then
        'WriteValue &H10, 0
        m_Enter = False
        tmrMOUSEOVER.Interval = 0
        RaiseEvent MouseLeave
        If Me.Gradient = True And m_ButtonStyle = True Then
            Me.Gradient = False
        End If
    End If
    
End Sub

'Public Function DrawLine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal oColor As OLE_COLOR = vbBlack, Optional ByVal Opacity As Integer = 100, Optional ByVal PenWidth As Integer = 1) As Boolean
'    Dim hGraphics As Long, hPen As Long
'
'    GdipCreateFromHDC hdc, hGraphics
'    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
'    GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth * nScale, UnitPixel, hPen
'    DrawLine = GdipDrawLineI(hGraphics, hPen, X1 * nScale, Y1 * nScale, X2 * nScale, Y2 * nScale) = 0
'    GdipDeletePen hPen
'    GdipDeleteGraphics hGraphics
'End Function
'
'
'Public Function Polygon(ByVal hdc As Long, ByVal PenWidth As Long, ByVal oColor As OLE_COLOR, ByVal Opacity As Integer, ParamArray vPoints() As Variant) As Boolean
'    Dim hGraphics As Long, hBrush As Long, hPen As Long
'    Dim lPoints() As Long
'    Dim lCount As Long
'    Dim i As Long
'
'    If UBound(vPoints) = 1 Then
'        lCount = vPoints(1)
'        ReDim lPoints(lCount - 1)
'        CopyMemory lPoints(0), ByVal CLng(vPoints(0)), lCount * 4
'    Else
'        lCount = UBound(vPoints) + 1
'        ReDim lPoints(lCount - 1)
'        For i = 0 To lCount - 1
'            lPoints(i) = vPoints(i) * nScale
'        Next
'    End If
'    GdipCreateFromHDC hdc, hGraphics
'    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
'
'    If PenWidth > 0 Then
'        GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth, UnitPixel, hPen
'        Call GdipDrawPolygonI(hGraphics, hPen, lPoints(0), lCount / 2)
'        GdipDeletePen hPen
'    Else
'        GdipCreateSolidFill ConvertColor(oColor, Opacity), hBrush
'        Call GdipFillPolygonI(hGraphics, hBrush, lPoints(0), lCount / 2, &H1)
'        GdipDeleteBrush hBrush
'    End If
'
'    GdipDeleteGraphics hGraphics
'End Function

Public Property Get IsMouseOver() As Boolean
    IsMouseOver = m_Over
End Property

Public Property Get IsMouseEnter() As Boolean
    IsMouseEnter = m_Enter
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
    If m_ButtonStyle Then
        If m_Invert Then
            Me.StatePressed = False
        Else
            Me.StatePressed = True
        End If
    End If
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, x, y)
    If m_ButtonStyle Then
        If m_Invert Then
            Me.StatePressed = False
        Else
            Me.StatePressed = True
        End If
        IsPressed = Not IsPressed
        'Me.Gradient = True
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    If m_ButtonStyle Then
        If Not m_ToggleButtonStyle Then
            If m_Invert Then
                Me.StatePressed = True
            Else
                Me.StatePressed = False
            End If
        Else
           ' Me.Gradient = False
            Me.StatePressed = IsPressed
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Debug.Print X, Y
    If m_ButtonStyle Then
        If IsPointInPath(x, y) Then
            If Me.Gradient = False Then
                Me.Gradient = True
                Me.GradientFlip = True
            End If
        Else
            If Me.Gradient = True Then
                Me.Gradient = False
            End If
        End If
    End If
    If hCur Then SetCursor hCur
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Function IsPointInPath(PX As Single, PY As Single) As Boolean
    Dim lResult As Long
    Dim x As Long, y As Long
    x = (m_Distance + m_Blur * 2) * nScale
    y = (m_Distance + m_Blur * 2) * nScale
    GdipIsVisiblePathPoint mPath, PX - x, PY - y, 0&, lResult
    IsPointInPath = lResult
End Function


Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(ByVal NewValue As IPictureDisp)
    UserControl.MouseIcon = NewValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MousePointerHands(ByVal NewValue As Boolean)
    m_MousePointerHands = NewValue
    If NewValue Then
        If Ambient.UserMode Then
            UserControl.MousePointer = vbCustom
            UserControl.MouseIcon = GetSystemHandCursor
        End If
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        UserControl.MousePointer = vbDefault
        UserControl.MouseIcon = Nothing
    End If
    PropertyChanged "MousePointerHands"
End Property

Public Property Get MousePointerHands() As Boolean
    MousePointerHands = m_MousePointerHands
End Property


Public Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, IPic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
 
    Set GetSystemHandCursor = IPic
    
End Function


