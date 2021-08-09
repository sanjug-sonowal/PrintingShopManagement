VERSION 5.00
Begin VB.UserControl ucScrollbar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
End
Attribute VB_Name = "ucScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucScrollbar.ctl
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  None
' Last revision: 12.20.2005
' Version:       1.0.4
'----------------------------------------------------------------------------------------
'
' (*) 1. Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'     2. pvCheckEnvironment() and pvIsLuna() routines by Paul Caton
'
'     3. Flat button fxs code extracted from (see pvDrawFlatButton() routine):
'        Special flat Cool Scrollbars version 1.2 by James Brown
'        http://www.catch22.net/tuts/coolscroll.asp
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.0.1: - Flat style *properly* painted:
'              * Hot thumb appearance = Pressed thumb appearance.
'              * Pressed/hot buttons using correct system colors.
'              Is there a default? For example, ListView with flat-scrollbars flag set,
'              preserves pressed buttons with 1-pixel edge using 'shadow' color and
'              their background is filled using color black instead of 'dark shadow'.
'   * 1.0.2: - Added Refresh method: only for custom-draw purposes.
'   * 1.0.3: - Fixed control on m_bHasTrack and m_bHasNullTrack flags.
'   * 1.0.4: - Fixed thumb rendering (classic style). DrawFrameControl->DrawEdge.
'----------------------------------------------------------------------------------------
'
' Notes:
'
'   * Restriction: Max >= Min
'   * Restriction: TabStop not supported
'----------------------------------------------------------------------------------------
'
' Known issues:
'========================================================================================

Option Explicit

Private Const VERSION_INFO As String = "1.0.3"

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                                           'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                                          'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE                          'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                         'Subclass data type
    hwnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private sc_aSubData()      As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED   As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC  As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04     As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05     As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08     As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09     As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'========================================================================================
' UserControl API declarations
'========================================================================================

Private Const SM_CXVSCROLL  As Long = 2
Private Const SM_CYHSCROLL  As Long = 3
Private Const SM_CYVSCROLL  As Long = 20
Private Const SM_CXHSCROLL  As Long = 21
Private Const SM_SWAPBUTTON As Long = 23

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Const SPI_GETKEYBOARDPREF  As Long = 68


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private Type RECT
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLUP       As Long = &H0
Private Const DFCS_SCROLLDOWN     As Long = &H1
Private Const DFCS_SCROLLLEFT     As Long = &H2
Private Const DFCS_SCROLLRIGHT    As Long = &H3
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000
Private Const DFCS_MONO           As Long = &H8000

Private Declare Function DrawFrameControl Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const BDR_RAISED As Long = &H5
Private Const BF_RECT    As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal Hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const COLOR_BTNFACE     As Long = 15
Private Const COLOR_3DSHADOW    As Long = 16
Private Const COLOR_BTNTEXT     As Long = 18
Private Const COLOR_3DHIGHLIGHT As Long = 20
Private Const COLOR_3DDKSHADOW  As Long = 21

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long

Private Const WHITE_BRUSH As Long = 0
Private Const BLACK_BRUSH As Long = 4

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
    
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long
 
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'//

Private Type PAINTSTRUCT
    Hdc             As Long
    fErase          As Long
    rcPaint         As RECT
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long

'---------
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWHEELSCROLLLINES As Long = 104
Private Const WHEEL_DELTA As Long = 120

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal Hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipDrawRectangle Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipFillRectangle Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipDrawLines Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTF, ByVal mCount As Long) As Long
Private Declare Function GdipFillPolygon Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByRef mPoints As POINTF, ByVal mCount As Long, ByVal mFillMode As FillMode) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As PenAlignment) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal mImage As Long) As Long


Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const GW_OWNER                  As Long = 4
Private Const WS_CHILD                  As Long = &H40000000
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Const LOGPIXELSX                As Long = 88
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Enum FillMode
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum

Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum
'//

Private Const WM_SIZE           As Long = &H5
Private Const WM_PAINT          As Long = &HF
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_CANCELMODE     As Long = &H1F
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_SETCURSOR      As Long = &H20

Private Const MK_LBUTTON        As Long = &H1
 
'//

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


' [UxThemeSCROLLBARParts]
Private Const SBP_ARROWBTN = 1
Private Const SBP_THUMBBTNHORZ = 2
Private Const SBP_THUMBBTNVERT = 3
Private Const SBP_LOWERTRACKHORZ = 4
Private Const SBP_UPPERTRACKHORZ = 5
Private Const SBP_LOWERTRACKVERT = 6
Private Const SBP_UPPERTRACKVERT = 7
Private Const SBP_GRIPPERHORZ = 8
Private Const SBP_GRIPPERVERT = 9
Private Const SBP_SIZEBOX = 10

' [UxThemeARROWBTNStates]
Private Const ABS_UPNORMAL = 1
Private Const ABS_UPHOT = 2
Private Const ABS_UPPRESSED = 3
Private Const ABS_UPDISABLED = 4
Private Const ABS_DOWNNORMAL = 5
Private Const ABS_DOWNHOT = 6
Private Const ABS_DOWNPRESSED = 7
Private Const ABS_DOWNDISABLED = 8
Private Const ABS_LEFTNORMAL = 9
Private Const ABS_LEFTHOT = 10
Private Const ABS_LEFTPRESSED = 11
Private Const ABS_LEFTDISABLED = 12
Private Const ABS_RIGHTNORMAL = 13
Private Const ABS_RIGHTHOT = 14
Private Const ABS_RIGHTPRESSED = 15
Private Const ABS_RIGHTDISABLED = 16

'========================================================================================
' UserControl enums., variables and constants
'========================================================================================

'-- Public enums.:

Public Enum sbOrientationCts
    [oVertical] = 0
    [oHorizontal] = 1
End Enum

Public Enum sbStyleCts
    [sClassic] = 0
    [sFlat] = 1
    [sThemed] = 2
    [sCustomDraw] = 3
End Enum

Public Enum sbOnPaintPartCts
    [ppTLButton] = 0
    [ppBRButton] = 1
    [ppTLTrack] = 2
    [ppBRTrack] = 3
    [ppNullTrack] = 4
    [ppThumb] = 5
End Enum

Public Enum sbOnPaintPartStateCts
    [ppsNormal] = 0
    [ppsPressed] = 1
    [ppsHot] = 2
    [ppsDisabled] = 3
End Enum

'-- Private enums.:

Private Enum eFlatButtonStateCts
    [fbsNormal] = 0
    [fbsSelected] = 1
    [fbsHot] = 2
End Enum

Public Enum eArrowStyle
    [Style1] = 0
    [Style2] = 1
    [Style3] = 2
    [Style4] = 3
    [Style5] = 4
End Enum

'-- Private constants:

Private Const HT_NOTHING          As Long = 0
Private Const HT_TLBUTTON         As Long = 1
Private Const HT_BRBUTTON         As Long = 2
Private Const HT_TLTRACK          As Long = 3
Private Const HT_BRTRACK          As Long = 4
Private Const HT_THUMB            As Long = 5

Private Const TIMERID_CHANGE1     As Long = 1
Private Const TIMERID_CHANGE2     As Long = 2
Private Const TIMERID_HOT         As Long = 3
Private Const TIMERID_LEAVE       As Long = 4
Private Const TIMERID_FADE        As Long = 5
Private Const TIMERID_AUTOHIDDE   As Long = 6

Private Const CHANGEDELAY_MIN     As Long = 0
Private Const CHANGEFREQUENCY_MIN As Long = 25
Private Const TIMERDT_HOT         As Long = 25


Private Const GRIPPERSIZE_MIN     As Long = 16

'-- Private variables:

Private m_bHasTrack               As Boolean
Private m_bHasNullTrack           As Boolean
Private m_uRctNullTrack           As RECT

Private m_uRctTLButton            As RECT
Private m_uRctBRButton            As RECT
Private m_uRctTLTrack             As RECT
Private m_uRctBRTrack             As RECT
Private m_uRctThumb               As RECT
Private m_lThumbOffset            As Long
Private m_uRctDrag                As RECT

Private m_bTLButtonPressed        As Boolean
Private m_bBRButtonPressed        As Boolean
Private m_bTLTrackPressed         As Boolean
Private m_bBRTrackPressed         As Boolean
Private m_bThumbPressed           As Boolean

Private m_bTLButtonHot            As Boolean
Private m_bBRButtonHot            As Boolean
Private m_bThumbHot               As Boolean

Private m_lAbsRange               As Long
Private m_lThumbPos               As Long
Private m_lThumbSize              As Long
Private m_eHitTest                As Long
Private m_eHitTestHot             As Long
Private m_x                       As Long
Private m_y                       As Long
Private m_lValueStartDrag         As Long

Private m_hPatternBrush           As Long

'-- Property variables:

Private m_lChangeDelay            As Long
Private m_lChangeFrequency        As Long
Private m_lMax                    As Long
Private m_lMin                    As Long
Private m_lValue                  As Long
Private m_lSmallChange            As Long
Private m_lLargeChange            As Long
Private m_eOrientation            As sbOrientationCts
Private m_eStyle                  As sbStyleCts
Private m_bShowButtons            As Boolean
Private m_RoundStyle               As Boolean
Private m_bIsXP                   As Boolean ' RO
Private m_bIsLuna                 As Boolean ' RO


Private mArrowStyle As eArrowStyle
'-- Default property values:

Private Const ENABLED_DEF         As Boolean = True
Private Const MIN_DEF             As Long = 0
Private Const MAX_DEF             As Long = 100
Private Const VALUE_DEF           As Long = MIN_DEF
Private Const SMALLCHANGE_DEF     As Long = 1
Private Const LARGECHANGE_DEF     As Long = 10
Private Const CHANGEDELAY_DEF     As Long = 500
Private Const CHANGEFREQUENCY_DEF As Long = 50
Private Const ORIENTATION_DEF     As Long = [oVertical]
Private Const STYLE_DEF           As Long = [sClassic]
Private Const SHOWBUTTONS_DEF     As Boolean = True

Private nScale As Single
Private Angle As Long
Private THUMBSIZE_MIN As Long
Private m_BackColor As OLE_COLOR
Private m_TrackColor As OLE_COLOR
Private m_ThemeColor As OLE_COLOR
Private m_FlatButtons As Boolean
Private m_FlatTrack As Boolean
Private m_MouseInControl As Boolean
Private m_AutoHidden As Boolean
Private m_Opacity As Integer
Private m_FadeIn As Boolean
Private m_TimeAutoHidde As Integer
Private m_HookContainer As Boolean
'-- Events:

Public Event Change()
Public Event Scroll()
Public Event ThemeChanged()
Public Event OnPaint(ByVal lHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal eState As sbOnPaintPartStateCts)

'//

'========================================================================================
' UserControl initialization/termination
'========================================================================================

Private Sub UserControl_Initialize()

    nScale = GetWindowsDPI
    
    THUMBSIZE_MIN = 16 * nScale
    
End Sub

Private Sub UserControl_Show()
    If m_AutoHidden And m_FadeIn = False Then
      m_Opacity = 0
      m_FadeIn = True
      pvSetTimer TIMERID_FADE, 10
    Else
        Refresh
    End If

End Sub

Private Sub UserControl_Terminate()
    
    On Error GoTo Catch
    
    '-- Stop subclassing
    Call Subclass_StopAll
    
Catch:
    On Error GoTo 0
  
    '-- In any case...

    Call pvKillTimer(TIMERID_FADE)
    Call pvKillTimer(TIMERID_LEAVE)
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
End Sub



'========================================================================================
' Only on design-mode
'========================================================================================

Private Sub UserControl_Resize()
    If (Ambient.UserMode = False) Then
        Call pvOnSize
    End If
    
End Sub

'Private Sub UserControl_Paint()
    'If (Ambient.UserMode = False) Then
        'Cls
    '    Call pvOnPaint(UserControl.HDC)
    'End If
'End Sub



'========================================================================================
' UserControl subclass procedure
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lhWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long _
                          )
Attribute zSubclass_Proc.VB_MemberFlags = "40"
                          
  Dim uPS As PAINTSTRUCT
  
    Select Case lhWnd
        
        Case UserControl.hwnd
        
            Select Case uMsg
            
                'Case WM_PAINT
                    'Call BeginPaint(lhWnd, uPS)
                    'Call pvOnPaint(uPS.hDC)
                    'Call EndPaint(lhWnd, uPS)
                    'bHandled = True: lReturn = 0
                    
                Case WM_SIZE
                    Call pvOnSize
                    bHandled = True: lReturn = 0
                
                Case WM_LBUTTONDOWN
                    Call pvOnMouseDown(wParam, lParam)
                    
                Case WM_MOUSEMOVE
                    Call pvOnMouseMove(wParam, lParam)
                
                Case WM_LBUTTONUP, WM_CANCELMODE
                    Call pvOnMouseUp
                   
                Case WM_LBUTTONDBLCLK
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    
                Case WM_TIMER
                    Call pvOnTimer(wParam)
                    
                Case WM_SYSCOLORCHANGE
                    Call pvOnSysColorChange
                    
                Case WM_MOUSEWHEEL
                    Dim ulScrollLines As Long
                    
                    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, ulScrollLines, 0
                
                    If wParam < 0 Then
                        Me.Value = m_lValue + (WHEEL_DELTA / ulScrollLines)
                    Else
                        Me.Value = m_lValue - (WHEEL_DELTA / ulScrollLines)
                    End If
            End Select
        Case Else
            Select Case uMsg
                Case WM_MOUSEWHEEL
                    If UserControl.Extender.Visible = False Then Exit Sub
                    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, ulScrollLines, 0
                
                    If wParam < 0 Then
                        Me.Value = m_lValue + Me.SmallChange '(WHEEL_DELTA / ulScrollLines) * 3
                    Else
                        Me.Value = m_lValue - Me.SmallChange  '(WHEEL_DELTA / ulScrollLines) * 3
                    End If
                    
                Case WM_SETCURSOR
                    Call pvKillTimer(TIMERID_AUTOHIDDE)
                    pvSetTimer TIMERID_AUTOHIDDE, m_TimeAutoHidde
                    If m_MouseInControl = False Then
                      m_MouseInControl = True
                      pvSetTimer TIMERID_LEAVE, 10
                      pvOnPaint UserControl.Hdc
                      
                      If m_AutoHidden And m_FadeIn = False Then
                        m_Opacity = 0
                        m_FadeIn = True
                        pvSetTimer TIMERID_FADE, 10
                      End If
                      
                    End If
            End Select
    End Select
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Sub Refresh()
    
    '-- Force a complete paint
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Sub



'========================================================================================
' Messages response
'========================================================================================

Private Sub pvOnSize()
 
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Sub



Private Sub pvOnMouseDown( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
  
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        
        Call pvMakePoints(lParam, m_x, m_y)
        m_eHitTest = pvHitTest(m_x, m_y)
        
        Select Case m_eHitTest
        
            Case HT_THUMB
                Select Case m_eOrientation
                    Case [oVertical]
                        m_lThumbOffset = m_uRctThumb.Y1 - m_y
                    Case [oHorizontal]
                        m_lThumbOffset = m_uRctThumb.X1 - m_x
                End Select
                m_bThumbPressed = True
                m_bThumbHot = False
                m_lValueStartDrag = m_lValue
                'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                pvOnPaint UserControl.Hdc
            Case HT_TLBUTTON
                m_bTLButtonPressed = True
                m_bTLButtonHot = False
                Call pvScrollPosDec(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_BRBUTTON
                m_bBRButtonPressed = True
                m_bBRButtonHot = False
                Call pvScrollPosInc(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_TLTRACK
                m_bTLTrackPressed = True
                Call pvScrollPosDec(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_BRTRACK
                m_bBRTrackPressed = True
                Call pvScrollPosInc(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
        End Select
    End If
End Sub

Private Sub pvOnMouseMove( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
  
  Dim lValuePrev As Long
  Dim lThumbPosPrev As Long
  Dim bPressed As Boolean
  Dim bHot As Boolean
  
    If m_MouseInControl = False Then
      m_MouseInControl = True
      pvSetTimer TIMERID_LEAVE, 10
      pvOnPaint UserControl.Hdc
      
      If m_AutoHidden And m_FadeIn = False Then
        m_Opacity = 0
        m_FadeIn = True
        pvSetTimer TIMERID_FADE, 10
      End If
      
    End If
    
    Call pvMakePoints(lParam, m_x, m_y)
                    
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        
        Select Case m_eHitTest
        
            Case HT_THUMB
            
                lValuePrev = m_lValue
                lThumbPosPrev = m_lThumbPos
                
                If (PtInRect(m_uRctDrag, m_x, m_y)) Then
                    
                    Select Case m_eOrientation
                        
                        Case [oVertical]
                        
                            m_lThumbPos = m_y + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.Y2) Then
                                m_lThumbPos = m_uRctTLButton.Y2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.Y1) Then
                                m_lThumbPos = m_uRctBRButton.Y1 - m_lThumbSize
                            End If
                        
                        Case [oHorizontal]
                        
                            m_lThumbPos = m_x + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.X2) Then
                                m_lThumbPos = m_uRctTLButton.X2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.X1) Then
                                m_lThumbPos = m_uRctBRButton.X1 - m_lThumbSize
                            End If
                    End Select
                    m_lValue = pvGetScrollPos()
                  
                  Else
                    
                    m_lValue = m_lValueStartDrag
                    m_lThumbPos = pvGetThumbPos()
                End If
                
                If (m_lThumbPos <> lThumbPosPrev) Then
                    Call pvSizeTrack
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    pvOnPaint UserControl.Hdc
                    If (m_lValue <> lValuePrev) Then
                        RaiseEvent Scroll
                    End If
                End If
            
            Case HT_TLBUTTON
                
                bPressed = (PtInRect(m_uRctTLButton, m_x, m_y) <> 0)
                If (bPressed Xor m_bTLButtonPressed) Then
                    m_bTLButtonPressed = bPressed
                    pvOnPaint UserControl.Hdc
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
                
            Case HT_BRBUTTON
                
                bPressed = (PtInRect(m_uRctBRButton, m_x, m_y) <> 0)
                If (bPressed Xor m_bBRButtonPressed) Then
                    m_bBRButtonPressed = bPressed
                    pvOnPaint UserControl.Hdc
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
        End Select
    
      Else
        
        m_eHitTestHot = pvHitTest(m_x, m_y)
        
        Select Case m_eHitTestHot
            
            Case HT_TLBUTTON
                bHot = (PtInRect(m_uRctTLButton, m_x, m_y) <> 0)
                If (m_bTLButtonHot Xor bHot) Then
                    m_bTLButtonHot = True
                    m_bBRButtonHot = False
                    m_bThumbHot = False
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    pvOnPaint UserControl.Hdc
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            
            Case HT_BRBUTTON
                bHot = (PtInRect(m_uRctBRButton, m_x, m_y) <> 0)
                If (m_bBRButtonHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = True
                    m_bThumbHot = False
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    pvOnPaint UserControl.Hdc
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            
            Case HT_THUMB
                
                bHot = (PtInRect(m_uRctThumb, m_x, m_y) <> 0)
                If (m_bThumbHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = False
                    m_bThumbHot = True
                    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    pvOnPaint UserControl.Hdc
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If

        End Select
    End If
End Sub

Private Sub pvOnMouseUp()

    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
    If (m_eHitTest = HT_THUMB) Then
        If (m_lValue <> m_lValueStartDrag) Then
            RaiseEvent Change
        End If
    End If
    m_eHitTest = HT_NOTHING
    
    m_bTLButtonPressed = False
    m_bBRButtonPressed = False
    m_bThumbPressed = False
    m_bTLTrackPressed = False
    m_bBRTrackPressed = False
    
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    pvOnPaint UserControl.Hdc
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnTimer(ByVal wParam As Long)
  
  Dim uPt As POINTAPI
  
    Select Case wParam
    
        Case TIMERID_CHANGE1
        
            Call pvKillTimer(TIMERID_CHANGE1)
            Call pvSetTimer(TIMERID_CHANGE2, m_lChangeFrequency)
       
        Case TIMERID_CHANGE2
        
            Select Case m_eHitTest
                
                Case HT_TLBUTTON
                    If (PtInRect(m_uRctTLButton, m_x, m_y)) Then
                        If (pvScrollPosDec(m_lSmallChange) = False) Then
                            Call pvKillTimer(TIMERID_CHANGE2)
                        End If
                    End If
                
                Case HT_BRBUTTON
                    If (PtInRect(m_uRctBRButton, m_x, m_y)) Then
                        If (pvScrollPosInc(m_lSmallChange) = False) Then
                            Call pvKillTimer(TIMERID_CHANGE2)
                        End If
                    End If
                    
                Case HT_TLTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos > m_y) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                              Else
                                m_bTLTrackPressed = False
                                'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                                pvOnPaint UserControl.Hdc
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos > m_x) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                              Else
                                m_bTLTrackPressed = False
                                'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                                pvOnPaint UserControl.Hdc
                            End If
                    End Select
                
                Case HT_BRTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos + m_lThumbSize < m_y) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                              Else
                                m_bBRTrackPressed = False
                                'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                                pvOnPaint UserControl.Hdc
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos + m_lThumbSize < m_x) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                              Else
                                m_bBRTrackPressed = False
                                'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                                pvOnPaint UserControl.Hdc
                            End If
                    End Select
           End Select
      
        Case TIMERID_HOT
            
            Call GetCursorPos(uPt)
            Call ScreenToClient(hwnd, uPt)
            
            Select Case True
                
                Case m_bTLButtonHot
                    If (PtInRect(m_uRctTLButton, uPt.X, uPt.Y) = 0) Then
                        m_bTLButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                        pvOnPaint UserControl.Hdc
                    End If
               
                Case m_bBRButtonHot
                    If (PtInRect(m_uRctBRButton, uPt.X, uPt.Y) = 0) Then
                        m_bBRButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                        pvOnPaint UserControl.Hdc
                    End If
               
                Case m_bThumbHot
                    If (PtInRect(m_uRctThumb, uPt.X, uPt.Y) = 0) Then
                        m_bThumbHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                        pvOnPaint UserControl.Hdc
                    End If
            End Select
        Case TIMERID_LEAVE
            If IsMouseInControl = False Then
                If Not GetKeyState(1) < 0 Then
                    m_MouseInControl = False
                    pvOnPaint UserControl.Hdc
                    Call pvKillTimer(TIMERID_LEAVE)
                    If m_AutoHidden Then
                        pvSetTimer TIMERID_AUTOHIDDE, m_TimeAutoHidde
                    End If
                End If
            End If
        Case TIMERID_FADE
            If m_FadeIn Then
                m_Opacity = m_Opacity + 5
                If m_Opacity >= 255 Then Call pvKillTimer(TIMERID_FADE): m_Opacity = 255
            Else
                m_Opacity = m_Opacity - 5
                If m_Opacity <= 0 Then Call pvKillTimer(TIMERID_FADE): m_FadeIn = False: m_Opacity = 0
            End If

            Me.Refresh
        Case TIMERID_AUTOHIDDE
            If IsMouseInControl = False Then
                m_FadeIn = False
                pvSetTimer TIMERID_FADE, 10
                Call pvKillTimer(TIMERID_AUTOHIDDE)
            Else
                Call pvKillTimer(TIMERID_AUTOHIDDE)
            End If
    End Select
End Sub

Private Sub pvOnSysColorChange()
    
    '-- Repaint all
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Sub

'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------

Private Sub pvSizeButtons()
 
 Dim uRct        As RECT
 Dim lButtonSize As Long
    
    Call GetClientRect(hwnd, uRct)
    m_bHasTrack = False
    m_bHasNullTrack = False
    
    Select Case m_eOrientation
        
        Case [oVertical]
        
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CYVSCROLL) * -CLng(m_bShowButtons)

            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .Y2) Then
                    If (2 * lButtonSize < .Y2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                        Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, 0, lButtonSize, .X2, .Y2 - lButtonSize)
                      Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, .Y2 \ 2)
                        Call SetRect(m_uRctBRButton, 0, .Y2 \ 2 + (.Y2 Mod 2), .X2, .Y2)
                        m_bHasNullTrack = CBool(.Y2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, 0, .Y2 \ 2, .X2, .Y2 \ 2 + 1)
                        End If
                    End If
                  Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                    Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                End If
            End With
            
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 250, 25)
            
        Case [oHorizontal]
            
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CXHSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .X2) Then
                    If (2 * lButtonSize < .X2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, lButtonSize, 0, .X2 - lButtonSize, .Y2)
                      Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2 \ 2, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 \ 2 + (.X2 Mod 2), 0, .X2, .Y2)
                        m_bHasNullTrack = CBool(.X2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, .X2 \ 2, 0, .X2 \ 2 + 1, .Y2)
                        End If
                    End If
                  Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                    Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                End If
            End With
            
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 25, 250)
    End Select
    
    '-- No track: avoid pvSizeTrack() calcs.
    If (m_bHasTrack = False) Then
        Call SetRectEmpty(m_uRctTLTrack)
        Call SetRectEmpty(m_uRctBRTrack)
        Call SetRectEmpty(m_uRctThumb)
    End If
End Sub

Private Sub pvSizeTrack()
 
    If (m_bHasTrack) Then
    
        '-- Tracks and thumbs exist
        Select Case m_eOrientation
            
            Case [oVertical]
                
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, 0, m_uRctTLButton.Y2, m_uRctTLButton.X2, m_lThumbPos)
                Call SetRect(m_uRctBRTrack, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.X2, m_uRctBRButton.Y1)
                Call SetRect(m_uRctThumb, 0, m_lThumbPos, m_uRctBRButton.X2, m_lThumbPos + m_lThumbSize)
                
            Case [oHorizontal]
            
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, m_uRctTLButton.X2, 0, m_lThumbPos, m_uRctTLButton.Y2)
                Call SetRect(m_uRctBRTrack, m_lThumbPos + m_lThumbSize, 0, m_uRctBRButton.X1, m_uRctBRButton.Y2)
                Call SetRect(m_uRctThumb, m_lThumbPos, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.Y2)
        End Select
    End If
End Sub

Private Function pvGetThumbSize() As Long
    
    On Error Resume Next
    
    Select Case m_eOrientation
        
        Case [oVertical]
        
            pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) \ (m_lAbsRange \ m_lLargeChange + 1)
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
            
        Case [oHorizontal]
        
            pvGetThumbSize = (m_uRctBRButton.X1 - m_uRctTLButton.X2) \ (m_lAbsRange \ m_lLargeChange + 1)
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
    End Select
    
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Controling value
'----------------------------------------------------------------------------------------

Private Function pvScrollPosDec( _
                 ByVal lSteps As Long, _
                 Optional ByVal bForceRepaint As Boolean = False _
                 ) As Boolean
    
  Dim bChange    As Boolean
  Dim lValuePrev As Long
        
    lValuePrev = m_lValue
    
    m_lValue = m_lValue - lSteps
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        pvOnPaint UserControl.Hdc
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    
    pvScrollPosDec = bChange
End Function

Private Function pvScrollPosInc( _
                 ByVal lSteps As Long, _
                 Optional ByVal bForceRepaint As Boolean = False _
                 ) As Boolean
    
  Dim bChange    As Boolean
  Dim lValuePrev As Long
        
    lValuePrev = m_lValue
    
    m_lValue = m_lValue + lSteps
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        pvOnPaint UserControl.Hdc
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    
    pvScrollPosInc = bChange
End Function

'----------------------------------------------------------------------------------------
' Positioning thumb and getting value from thumb position
'----------------------------------------------------------------------------------------

Private Function pvGetThumbPos() As Long

    On Error Resume Next
    
    Select Case m_eOrientation
        Case [oVertical]
            pvGetThumbPos = m_uRctTLButton.Y2
            pvGetThumbPos = pvGetThumbPos + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
        Case [oHorizontal]
            pvGetThumbPos = m_uRctTLButton.X2
            pvGetThumbPos = pvGetThumbPos + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
    End Select
    
    On Error GoTo 0
End Function

Private Function pvGetScrollPos() As Long
    
    On Error Resume Next
    
    Select Case m_eOrientation
        Case [oVertical]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.Y2) / (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) * m_lAbsRange
        Case [oHorizontal]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.X2) / (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) * m_lAbsRange
    End Select
    
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Hit-Test
'----------------------------------------------------------------------------------------

Private Function pvHitTest(ByVal X As Long, ByVal Y As Long) As Long
    
    Select Case True
        Case PtInRect(m_uRctTLButton, X, Y)
            pvHitTest = HT_TLBUTTON
        Case PtInRect(m_uRctBRButton, X, Y)
            pvHitTest = HT_BRBUTTON
        Case PtInRect(m_uRctTLTrack, X, Y)
            pvHitTest = HT_TLTRACK
        Case PtInRect(m_uRctBRTrack, X, Y)
            pvHitTest = HT_BRTRACK
        Case PtInRect(m_uRctThumb, X, Y)
            pvHitTest = HT_THUMB
    End Select
End Function

Private Sub pvMakePoints( _
            ByVal lPoint As Long, _
            X As Long, _
            Y As Long _
            )
            
    If (lPoint And &H8000&) Then
        X = &H8000 Or (lPoint And &H7FFF&)
      Else
        X = lPoint And &HFFFF&
    End If
    If (lPoint And &H80000000) Then
        Y = (lPoint \ &H10000) - 1
      Else
        Y = lPoint \ &H10000
    End If
End Sub

'----------------------------------------------------------------------------------------
' Timing
'----------------------------------------------------------------------------------------

Private Sub pvSetTimer( _
            ByVal lTimerID As Long, _
            ByVal ldT As Long _
            )
    
    Call SetTimer(UserControl.hwnd, lTimerID, ldT, 0)
End Sub

Private Sub pvKillTimer( _
            ByVal lTimerID As Long _
            )
            
    Call KillTimer(UserControl.hwnd, lTimerID)
    m_eHitTestHot = HT_NOTHING
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------

Private Sub pvDrawFlatButton( _
            ByVal Hdc As Long, _
            uRct As RECT, _
            ByVal lfArrowDirection As Long, _
            ByVal eState As eFlatButtonStateCts _
            )

  Dim uRctMem    As RECT
  
  Dim hDCMem1    As Long
  Dim hDCMem2    As Long
  Dim hBmp1      As Long
  Dim hBmp2      As Long
  Dim hBmpOld1   As Long
  Dim hBmpOld2   As Long
  
  Dim clrBkOld   As Long
  Dim clrTextOld As Long
        
    With uRct
    
        '-- Monochrome bitmap to convert the arrow to black/white mask
        hDCMem1 = CreateCompatibleDC(Hdc)
        hBmp1 = CreateBitmap(.X2 - .X1, .Y2 - .Y1, 1, 1, ByVal 0)
        hBmpOld1 = SelectObject(hDCMem1, hBmp1)
        
        '-- Normal bitmap to draw the arrow into
        hDCMem2 = CreateCompatibleDC(Hdc)
        hBmp2 = CreateCompatibleBitmap(Hdc, .X2 - .X1, .Y2 - .Y1)
        hBmpOld2 = SelectObject(hDCMem2, hBmp2)
        
        '-- Draw frame normaly
        Call CopyRect(uRctMem, uRct)
        Call OffsetRect(uRctMem, -.X1, -.Y1)
        Call DrawFrameControl(hDCMem2, uRctMem, DFC_SCROLL, DFCS_FLAT Or lfArrowDirection)
        
        Select Case eState
        
            Case [fbsNormal]
                
                '-- Nothing to do
                Call BitBlt(Hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            
            Case [fbsSelected]
                
                '-- Invert
                Call InvertRect(hDCMem2, uRctMem)
                Call BitBlt(Hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            
            Case [fbsHot]
            
                '-- Mask glyph
                Call SetBkColor(hDCMem2, GetSysColor(COLOR_BTNTEXT))
                Call BitBlt(hDCMem1, 0, 0, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
                clrBkOld = SetBkColor(Hdc, GetSysColor(COLOR_3DHIGHLIGHT))
                clrTextOld = SetTextColor(Hdc, GetSysColor(COLOR_3DSHADOW))
                Call BitBlt(Hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem1, 0, 0, vbSrcCopy)
                Call SetBkColor(Hdc, clrBkOld)
                Call SetTextColor(Hdc, clrTextOld)
        End Select
    End With
        
    '-- Clean up
    Call DeleteObject(SelectObject(hDCMem1, hBmpOld1))
    Call DeleteObject(SelectObject(hDCMem2, hBmpOld2))
    Call DeleteDC(hDCMem1)
    Call DeleteDC(hDCMem2)
End Sub

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------



'-- Checking environment and Luna theming

Private Sub pvCheckEnvironment()

  Dim uOSV As OSVERSIONINFO
    
    m_bIsXP = False
    m_bIsLuna = False
    
    With uOSV
        
        .dwOSVersionInfoSize = Len(uOSV)
        Call GetVersionEx(uOSV)
        
        If (.dwPlatformId = 2) Then
            If (.dwMajorVersion = 5) Then     ' NT based
                If (.dwMinorVersion > 0) Then ' XP
                    m_bIsXP = True
                    
                End If
            End If
        End If
    End With
End Sub





'========================================================================================
' UserControl persistent properties
'========================================================================================

Private Sub UserControl_InitProperties()
    
    '-- Initialization default values
    ManageGDIToken UserControl.ContainerHwnd
    
    Let m_lChangeDelay = CHANGEDELAY_DEF
    Let m_lChangeFrequency = CHANGEFREQUENCY_DEF
    Let m_lMin = MIN_DEF
    Let m_lMax = MAX_DEF
    Let m_lValue = VALUE_DEF
    Let m_lSmallChange = SMALLCHANGE_DEF
    Let m_lLargeChange = LARGECHANGE_DEF
    Let m_eOrientation = ORIENTATION_DEF
    Let m_eStyle = STYLE_DEF
    Let m_bShowButtons = SHOWBUTTONS_DEF
    m_BackColor = Ambient.BackColor
    m_TrackColor = vbWindowBackground
    m_ThemeColor = vbScrollBars
    m_FlatButtons = False
    m_FlatTrack = False
    m_TimeAutoHidde = 2000
    m_Opacity = 255
    m_RoundStyle = True
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ManageGDIToken UserControl.ContainerHwnd
    '-- Bag properties
    With PropBag
        
        '-- Read inherently-stored properties
        Let UserControl.Enabled = .ReadProperty("Enabled", ENABLED_DEF)
        
        '-- Read 'memory' properties
        Let m_lMin = .ReadProperty("Min", MIN_DEF)
        Let m_lMax = .ReadProperty("Max", MAX_DEF)
        Let m_lValue = .ReadProperty("Value", VALUE_DEF)
        Let m_lSmallChange = .ReadProperty("SmallChange", SMALLCHANGE_DEF)
        Let m_lLargeChange = .ReadProperty("LargeChange", LARGECHANGE_DEF)
        Let m_lChangeDelay = .ReadProperty("ChangeDelay", CHANGEDELAY_DEF)
        Let m_lChangeFrequency = .ReadProperty("ChangeFrequency", CHANGEFREQUENCY_DEF)
        Let m_eOrientation = .ReadProperty("Orientation", ORIENTATION_DEF)
        Let m_eStyle = .ReadProperty("Style", STYLE_DEF)
        Let m_bShowButtons = .ReadProperty("ShowButtons", SHOWBUTTONS_DEF)
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_TrackColor = .ReadProperty("TrackColor", vbWindowBackground)
        m_ThemeColor = .ReadProperty("ThemeColor", vbScrollBars)
        m_FlatButtons = .ReadProperty("FlatButtons", False)
        m_FlatTrack = .ReadProperty("FlatTrack", False)
        m_AutoHidden = .ReadProperty("AutoHidden", False)
        m_TimeAutoHidde = .ReadProperty("TimeAutoHidde", 2000)
        m_RoundStyle = .ReadProperty("RoundStyle", True)
        m_HookContainer = .ReadProperty("HookContainer", False)
    End With
    If m_AutoHidden Then
        m_Opacity = 0
    Else
        m_Opacity = 255
    End If
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call pvOnPaint(UserControl.Hdc)
    '-- Run-time?
    If (Ambient.UserMode) Then
        
        '-- Check OS and Luna theme
        Call pvCheckEnvironment
    
        '-- Subclass UC window and process following messages
        Call Subclass_Start(UserControl.hwnd)
        'Call Subclass_AddMsg(UserControl.hwnd, WM_PAINT, [MSG_BEFORE])
        Call Subclass_AddMsg(UserControl.hwnd, WM_SIZE, [MSG_BEFORE])
        Call Subclass_AddMsg(UserControl.hwnd, WM_CANCELMODE)
        Call Subclass_AddMsg(UserControl.hwnd, WM_MOUSEMOVE)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONDOWN)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONUP)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONDBLCLK)
        Call Subclass_AddMsg(UserControl.hwnd, WM_TIMER)
        Call Subclass_AddMsg(UserControl.hwnd, WM_SYSCOLORCHANGE)
        If m_HookContainer Then
            Call Subclass_AddMsg(UserControl.hwnd, WM_MOUSEWHEEL)
        End If
        If (m_bIsXP) Then
            Call Subclass_AddMsg(UserControl.hwnd, WM_THEMECHANGED)
        End If
        
        
        If m_HookContainer Then
            Call Subclass_Start(UserControl.ContainerHwnd)
            Call Subclass_AddMsg(UserControl.ContainerHwnd, WM_MOUSEWHEEL)
            Call Subclass_AddMsg(UserControl.ContainerHwnd, WM_SETCURSOR)
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, ENABLED_DEF)
        Call .WriteProperty("Min", m_lMin, MIN_DEF)
        Call .WriteProperty("Max", m_lMax, MAX_DEF)
        Call .WriteProperty("Value", m_lValue, VALUE_DEF)
        Call .WriteProperty("SmallChange", m_lSmallChange, SMALLCHANGE_DEF)
        Call .WriteProperty("LargeChange", m_lLargeChange, LARGECHANGE_DEF)
        Call .WriteProperty("ChangeDelay", m_lChangeDelay, CHANGEDELAY_DEF)
        Call .WriteProperty("ChangeFrequency", m_lChangeFrequency, CHANGEFREQUENCY_DEF)
        Call .WriteProperty("Orientation", m_eOrientation, ORIENTATION_DEF)
        Call .WriteProperty("Style", m_eStyle, STYLE_DEF)
        Call .WriteProperty("ShowButtons", m_bShowButtons, SHOWBUTTONS_DEF)
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("TrackColor", m_TrackColor, vbWindowBackground)
        Call .WriteProperty("ThemeColor", m_ThemeColor, vbScrollBars)
        Call .WriteProperty("FlatButtons", m_FlatButtons, False)
        Call .WriteProperty("FlatTrack", m_FlatTrack, False)
        Call .WriteProperty("AutoHidden", m_AutoHidden, False)
        Call .WriteProperty("RoundStyle", m_RoundStyle, True)
        Call .WriteProperty("HookContainer", m_HookContainer, False)
        
    End With
End Sub


Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
    PropertyChanged "BackColor"

    Refresh
End Property

Public Property Get TrackColor() As OLE_COLOR
    TrackColor = m_TrackColor
End Property

Public Property Let TrackColor(ByVal New_Value As OLE_COLOR)
    m_TrackColor = New_Value
    PropertyChanged "TrackColor"

    Refresh
End Property

Public Property Get ThemeColor() As OLE_COLOR
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal New_Value As OLE_COLOR)
    m_ThemeColor = New_Value
    PropertyChanged "ThemeColor"

    Refresh
End Property

Public Property Get FlatButtons() As Boolean
    FlatButtons = m_FlatButtons
End Property

Public Property Let FlatButtons(ByVal New_Value As Boolean)
    m_FlatButtons = New_Value
    PropertyChanged "FlatButtons"

    Refresh
End Property

Public Property Get FlatTrack() As Boolean
    FlatTrack = m_FlatTrack
End Property

Public Property Let FlatTrack(ByVal New_Value As Boolean)
    m_FlatTrack = New_Value
    PropertyChanged "FlatTrack"

    Refresh
End Property

Public Property Get AutoHidden() As Boolean
    AutoHidden = m_AutoHidden
End Property

Public Property Let AutoHidden(ByVal New_Value As Boolean)
    m_AutoHidden = New_Value
    PropertyChanged "AutoHidden"

    Refresh
End Property

Public Property Get TimeAutoHidde() As Integer
    TimeAutoHidde = m_TimeAutoHidde
End Property

Public Property Let TimeAutoHidde(ByVal New_Value As Integer)
    m_TimeAutoHidde = New_Value
    PropertyChanged "TimeAutoHidde"

    Refresh
End Property

Public Property Get RoundStyle() As Boolean
    RoundStyle = m_RoundStyle
End Property

Public Property Let RoundStyle(ByVal New_Value As Boolean)
    m_RoundStyle = New_Value
    PropertyChanged "RoundStyle"

    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enable As Boolean)
    UserControl.Enabled = New_Enable
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property

Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < m_lMin) Then
        New_Max = m_lMin
    End If
    m_lMax = New_Max
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get Min() As Long
    Min = m_lMin
End Property

Public Property Let Min(ByVal New_Min As Long)
    If (New_Min > m_lMax) Then
        New_Min = m_lMax
    End If
    m_lMin = New_Min
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = m_lValue
End Property

Public Property Let Value(ByVal New_Value As Long)

  Dim lValuePrev As Long

    If (New_Value < m_lMin) Then
        New_Value = m_lMin
    ElseIf (New_Value > m_lMax) Then
        New_Value = m_lMax
    End If
    lValuePrev = m_lValue
    m_lValue = New_Value
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    
      If m_AutoHidden And m_FadeIn = False Then
       ' m_Opacity = 0
       ' m_FadeIn = True
       ' pvSetTimer TIMERID_FADE, 10
      End If
    
    pvOnPaint UserControl.Hdc
    If (m_lValue <> lValuePrev) Then
        RaiseEvent Change
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_lSmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Long)
    If (New_SmallChange < 1) Then
        New_SmallChange = 1
    End If
    m_lSmallChange = New_SmallChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_lLargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    If (New_LargeChange < 1) Then
        New_LargeChange = 1
    End If
    m_lLargeChange = New_LargeChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get ChangeDelay() As Long
    ChangeDelay = m_lChangeDelay
End Property

Public Property Let ChangeDelay(ByVal New_ChangeDelay As Long)
    If (New_ChangeDelay < CHANGEDELAY_MIN) Then
        New_ChangeDelay = CHANGEDELAY_MIN
    End If
    m_lChangeDelay = New_ChangeDelay
End Property

Public Property Get ChangeFrequency() As Long
    ChangeFrequency = m_lChangeFrequency
End Property

Public Property Let ChangeFrequency(ByVal New_ChangeFrequency As Long)
    If (New_ChangeFrequency < CHANGEFREQUENCY_MIN) Then
        New_ChangeFrequency = CHANGEFREQUENCY_MIN
    End If
    m_lChangeFrequency = New_ChangeFrequency
End Property

Public Property Get Orientation() As sbOrientationCts
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As sbOrientationCts)
    If (New_Orientation < [oVertical]) Then
        New_Orientation = [oVertical]
    ElseIf (New_Orientation > [oHorizontal]) Then
        New_Orientation = [oHorizontal]
    End If
    m_eOrientation = New_Orientation
    Call pvOnSize
End Property

Public Property Get Style() As sbStyleCts
    Style = m_eStyle
End Property

Public Property Let Style(ByVal New_Style As sbStyleCts)
    If (New_Style < [sClassic]) Then
        New_Style = [sClassic]
    ElseIf (New_Style > [sCustomDraw]) Then
        New_Style = [sCustomDraw]
    End If
    m_eStyle = New_Style
    'Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    pvOnPaint UserControl.Hdc
End Property

Public Property Get ShowButtons() As Boolean
    ShowButtons = m_bShowButtons
End Property

Public Property Let ShowButtons(ByVal New_ShowButtons As Boolean)
    m_bShowButtons = New_ShowButtons
    Call pvOnSize
End Property

'// Runtime read only

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get IsXP() As Boolean
    IsXP = m_bIsXP
End Property

Public Property Get IsThemed() As Boolean
    IsThemed = m_bIsLuna
End Property

Public Property Get HookContainer() As Boolean
    HookContainer = m_HookContainer
End Property

Public Property Let HookContainer(ByVal New_Value As Boolean)
    
    
    If New_Value Then
        Call Subclass_Start(UserControl.ContainerHwnd)
        Call Subclass_AddMsg(UserControl.ContainerHwnd, WM_MOUSEWHEEL)
        Call Subclass_AddMsg(UserControl.ContainerHwnd, WM_SETCURSOR)
    Else
        If m_HookContainer Then
            Call Subclass_Stop(UserControl.ContainerHwnd)
        End If
    End If
    m_HookContainer = New_Value
    PropertyChanged "HookContainer"
End Property


'========================================================================================
' About
'========================================================================================

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    Call VBA.MsgBox("ucScrollbar " & VERSION_INFO & " - Carles P.V. 2005", , "About")
End Sub



'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  
    With sc_aSubData(zIdx(lhWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lhWnd As Long) As Long

  Const CODE_LEN              As Long = 202
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim j                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
    If (aBuf(1) = 0) Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        If (Subclass_InIDE) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lhWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lhWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        Call CopyMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_Stop(ByVal lhWnd As Long)
  
    With sc_aSubData(zIdx(lhWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB()
        Erase .aMsgTblA()
    End With
End Sub

Private Sub Subclass_StopAll()
  
  Dim i As Long
  
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hwnd <> 0) Then
                Call Subclass_Stop(.hwnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
  
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
      Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hwnd = lhWnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hwnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

' funcion para convertir un color long a un BGRA(Blue, Green, Red, Alpha)
Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function


Public Function ShiftColor(ByVal Color As Long, ByVal d As Long) As Long
    Dim BGRA(0 To 3) As Byte
    Dim R As Long, B As Long, G As Long
    
    OleTranslateColor Color, 0, VarPtr(Color)
    
    R = (Color And &HFF) + d
    G = ((Color \ &H100) Mod &H100) + d
    B = ((Color \ &H10000) Mod &H100) + d
    
    If (d > 0) Then
        If (R > &HFF) Then R = &HFF
        If (G > &HFF) Then G = &HFF
        If (B > &HFF) Then B = &HFF
    ElseIf (d < 0) Then
        If (R < 0) Then R = 0
        If (G < 0) Then G = 0
        If (B < 0) Then B = 0
    End If
    
    
    BGRA(0) = B
    BGRA(1) = G
    BGRA(2) = R
    BGRA(3) = m_Opacity
    
    CopyMemory ShiftColor, BGRA(0), 4&

End Function

Private Sub pvOnPaint(ByVal lHdc As Long)
    Dim hGraphics As Long
    Dim hBrush As Long
    Dim hImage As Long
    
    'UserControl.Cls
    Angle = THUMBSIZE_MIN
    If m_eOrientation = oVertical Then
       If Angle > UserControl.ScaleWidth Then Angle = UserControl.ScaleWidth
    Else
         If Angle > UserControl.ScaleHeight Then Angle = UserControl.ScaleHeight
    End If
  
    'GdipCreateBitmapFromScan0 UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    'GdipGetImageGraphicsContext hImage, hGraphics
    GdipCreateFromHDC Hdc, hGraphics
    
    
    
    GdipCreateSolidFill ConvertColor(m_BackColor, 100), hBrush
    GdipFillRectangle hGraphics, hBrush, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    GdipDeleteBrush hBrush
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    DrawTrack hGraphics
    

            
    If m_bShowButtons Then
        With m_uRctTLButton
            DrawButton hGraphics, .X1, .Y1, .X2, .Y2, True
        End With
        
        With m_uRctBRButton
            DrawButton hGraphics, .X1, .Y1, .X2, .Y2, False
        End With
    End If
    
    
    If (m_bHasTrack) Then
        With m_uRctTLTrack
            'RaiseEvent OnPaint(lhDC, .x1, .y1, .x2, .y2, [ppTLTrack], IIf(m_bTLTrackPressed, [ppsPressed], [ppsNormal]))
        End With
        With m_uRctBRTrack
            'RaiseEvent OnPaint(lhDC, .x1, .y1, .x2, .y2, [ppBRTrack], IIf(m_bBRTrackPressed, [ppsPressed], [ppsNormal]))
        End With
        With m_uRctThumb
            'RaiseEvent OnPaint(lhDC, .x1, .y1, .x2, .y2, [ppThumb], IIf(m_bThumbHot, [ppsHot], IIf(m_bThumbPressed, [ppsPressed], [ppsNormal])))
            DrawThumb hGraphics, .X1, .Y1, .X2, .Y2
        End With
    End If
    If (m_bHasNullTrack) Then
        With m_uRctNullTrack
            'RaiseEvent OnPaint(lhDC, .x1, .y1, .x2, .y2, [ppNullTrack], [ppsNormal])
        End With
    End If
    
    GdipDeleteGraphics hGraphics
    'GdipCreateFromHDC HDC, hGraphics
    'GdipDrawImageRectI hGraphics, hImage, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    'GdipDisposeImage hImage
    'GdipDeleteGraphics hGraphics
    UserControl.Refresh
End Sub

Private Function IsLightColor(Color As Long) As Boolean
    Dim R As Integer, G As Integer, B As Integer
    R = &HFF& And Color
    G = (&HFF00& And Color) \ 256
    B = (&HFF0000 And Color) \ 65536
    If (R + G + B) / 3 > 127 Then IsLightColor = True
End Function

Private Sub DrawButton(hGraphics As Long, X As Long, Y As Long, W As Long, H As Long, bFirst As Boolean)
    Dim hBrush As Long, hPen As Long
    Dim mPath As Long
    Dim lColor As Long
    Dim Width As Long
    Dim Height As Long
    Dim P As Long
    
    If IsLightColor(m_ThemeColor) Then P = -1 Else P = 1
    
    If m_FlatButtons Then
        If bFirst And Not m_bTLButtonHot And Not m_bTLButtonPressed Then
            DrawArrow hGraphics, X, Y, W - X, H - Y, bFirst
            Exit Sub
        End If
        If Not bFirst And Not m_bBRButtonHot And Not m_bBRButtonPressed Then
            DrawArrow hGraphics, X, Y, W - X, H - Y, bFirst
            Exit Sub
        End If
    End If
    
    Width = W - X - 1
    Height = H - Y - 1
    
    
    lColor = ShiftColor(m_ThemeColor, P * 20)
    
    
    If bFirst Then
        If m_bTLButtonHot Then lColor = ShiftColor(m_ThemeColor, 0)
        If m_bTLButtonPressed Then lColor = ShiftColor(m_ThemeColor, P * 20)
    Else
        If m_bBRButtonHot Then lColor = ShiftColor(m_ThemeColor, P * 20)
        If m_bBRButtonPressed Then lColor = ShiftColor(m_ThemeColor, P * 40)
    End If
    
    If Not m_RoundStyle Then
        GdipCreateSolidFill lColor, hBrush
        GdipFillRectangle hGraphics, hBrush, X, Y, Width, Height
        GdipDeleteBrush hBrush
        
        GdipCreatePen1 ShiftColor(m_ThemeColor, P * 40), 1, &H2, hPen
        GdipDrawRectangle hGraphics, hPen, X, Y, Width, Height
        GdipDeletePen hPen
    
        DrawArrow hGraphics, X, Y, W - X, H - Y, bFirst
        Exit Sub
    End If
    
    If GdipCreatePath(&H0, mPath) = 0 Then
    
        If m_eOrientation = oHorizontal Then
            If bFirst Then
                GdipAddPathArcI mPath, X, Y, Angle, Angle, 180, 90
                GdipAddPathArcI mPath, X + Width, Y, Angle, Angle, -90, -90
                GdipAddPathArcI mPath, X + Width, Y + Height - Angle, Angle, Angle, -180, -90
                GdipAddPathArcI mPath, X, Y + Height - Angle, Angle, Angle, 90, 90
            Else
                GdipAddPathArcI mPath, X - Angle, Y, Angle, Angle, 0, -90
                GdipAddPathArcI mPath, X + Width - Angle, Y, Angle, Angle, -90, 90
                GdipAddPathArcI mPath, X + Width - Angle, Y + Height - Angle, Angle, Angle, 0, 90
                GdipAddPathArcI mPath, X - Angle, Y + Height - Angle, Angle, Angle, 90, -90
            End If
        Else
            If bFirst Then
                GdipAddPathArcI mPath, X, Y, Angle, Angle, 180, 90
                GdipAddPathArcI mPath, X + Width - Angle, Y, Angle, Angle, 270, 90
                GdipAddPathArcI mPath, X + Width - Angle, Y + Height, Angle, Angle, 0, -90
                GdipAddPathArcI mPath, X, Y + Height, Angle, Angle, -90, -90
            Else
                GdipAddPathArcI mPath, X, Y - Angle, Angle, Angle, -180, -90
                GdipAddPathArcI mPath, X + Width - Angle, Y - Angle, Angle, Angle, -270, -90
                GdipAddPathArcI mPath, X + Width - Angle, Y + Height - Angle, Angle, Angle, 0, 90
                GdipAddPathArcI mPath, X, Y + Height - Angle, Angle, Angle, 90, 90
            End If
        End If
        
        GdipClosePathFigures mPath

        GdipCreateSolidFill lColor, hBrush
        GdipFillPath hGraphics, hBrush, mPath
        GdipDeleteBrush hBrush
 
        GdipCreatePen1 ShiftColor(m_ThemeColor, P * 40), 1, &H2, hPen
        GdipDrawPath hGraphics, hPen, mPath
        GdipDeletePen hPen
        
        GdipDeletePath mPath
        
    End If

    DrawArrow hGraphics, X, Y, W - X, H - Y, bFirst
End Sub

Private Sub DrawTrack(hGraphics As Long)
    Dim hBrush As Long, hPen As Long
    Dim mPath As Long
    Dim Width As Long
    Dim Height As Long
    Dim P As Long
    
    If IsLightColor(m_ThemeColor) Then P = -1 Else P = 1
    
    If m_FlatTrack Then
        'If Not m_MouseInControl Then Exit Sub
        Exit Sub
    End If
    
    Width = UserControl.ScaleWidth - 1
    Height = UserControl.ScaleHeight - 1
    
    If Not m_RoundStyle Then
        GdipCreateSolidFill ShiftColor(m_TrackColor, 0), hBrush
        GdipFillRectangle hGraphics, hBrush, 0, 0, Width, Height
        GdipDeleteBrush hBrush
        
        GdipCreatePen1 ShiftColor(m_TrackColor, -40), 1, &H2, hPen
        GdipDrawRectangle hGraphics, hPen, 0, 0, Width, Height
        GdipDeletePen hPen
        Exit Sub
    End If

    If GdipCreatePath(&H0, mPath) = 0 Then
        GdipAddPathArcI mPath, 0, 0, Angle, Angle, 180, 90
        GdipAddPathArcI mPath, Width - Angle, 0, Angle, Angle, 270, 90
        GdipAddPathArcI mPath, Width - Angle, Height - Angle, Angle, Angle, 0, 90
        GdipAddPathArcI mPath, 0, Height - Angle, Angle, Angle, 90, 90
        GdipClosePathFigures mPath
    
    
        GdipCreateSolidFill ShiftColor(m_TrackColor, 0), hBrush
        GdipFillPath hGraphics, hBrush, mPath
        GdipDeleteBrush hBrush
 
        GdipCreatePen1 ShiftColor(m_TrackColor, P * 20), 1, &H2, hPen
        'GdipSetPenMode hPen, PenAlignmentInset
        GdipDrawPath hGraphics, hPen, mPath
        GdipDeletePen hPen
        
        GdipDeletePath mPath
        
    End If
End Sub

Private Sub DrawThumb(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
    Dim hBrush As Long, hPen As Long
    Dim mPath As Long
    Dim Width As Long
    Dim Height As Long
    Dim lColor As Long
    Dim P As Long
    
    If IsLightColor(m_ThemeColor) Then P = -1 Else P = 1
    
    Width = W - X - 1
    Height = H - Y - 1
    
  
    lColor = ShiftColor(m_ThemeColor, 0)
    If m_bThumbHot Then lColor = ShiftColor(m_ThemeColor, P * 20)
    If m_bThumbPressed Then lColor = ShiftColor(m_ThemeColor, P * 40)

    If Not m_RoundStyle Then
        GdipCreateSolidFill lColor, hBrush
        GdipFillRectangle hGraphics, hBrush, X, Y, Width, Height
        GdipDeleteBrush hBrush
        
        GdipCreatePen1 ShiftColor(m_ThemeColor, P * 40), 1, &H2, hPen
        GdipDrawRectangle hGraphics, hPen, X, Y, Width, Height
        GdipDeletePen hPen
        Exit Sub
    End If

    If GdipCreatePath(&H0, mPath) = 0 Then
        GdipAddPathArcI mPath, X, Y, Angle, Angle, 180, 90
        GdipAddPathArcI mPath, X + Width - Angle, Y, Angle, Angle, 270, 90
        GdipAddPathArcI mPath, X + Width - Angle, Y + Height - Angle, Angle, Angle, 0, 90
        GdipAddPathArcI mPath, X, Y + Height - Angle, Angle, Angle, 90, 90
        GdipClosePathFigures mPath



        GdipCreateSolidFill lColor, hBrush
        GdipFillPath hGraphics, hBrush, mPath
        GdipDeleteBrush hBrush
 
        GdipCreatePen1 ShiftColor(m_ThemeColor, P * 40), 1, &H2, hPen
        'GdipSetPenMode hPen, PenAlignmentInset
        GdipDrawPath hGraphics, hPen, mPath
        GdipDeletePen hPen
        
        GdipDeletePath mPath
        
    End If
    
    'DrawArrow hGraphics, X, Y, W - X, H - Y, bFirst
    
End Sub



Private Sub DrawArrow(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Top As Boolean)
    Dim hBrush As Long, hPen As Long
    Dim PT() As POINTF
    
    mArrowStyle = Style3
        
    If m_eOrientation = oVertical Then
        X = X + W / 2 - H / 2
    Else
        Y = X
        X = H / 2 - W / 2
        H = W
    End If
    
    Select Case mArrowStyle
    
        Case Style1
            ReDim PT(5)
            GdipCreateSolidFill ConvertColor(vbBlack, 100), hBrush
            If Top Then
                PT(0).X = X + H / 3:  PT(0).Y = Y + H / 1.5
                PT(1).X = X + H / 2:  PT(1).Y = Y + H / 2
                PT(2).X = X + H / 1.5: PT(2).Y = Y + H / 1.5
                PT(3).X = X + H / 1.5: PT(3).Y = Y + H / 1.8
                PT(4).X = X + H / 2:  PT(4).Y = Y + H / 2.7
                PT(5).X = X + H / 3:  PT(5).Y = Y + H / 1.8
            Else
                PT(0).X = X + H / 3:    PT(0).Y = Y + H / 3
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 2
                PT(2).X = X + H / 1.5:  PT(2).Y = Y + H / 3
                PT(3).X = X + H / 1.5:  PT(3).Y = Y + H / 2.1
                PT(4).X = X + H / 2:    PT(4).Y = Y + H / 1.6
                PT(5).X = X + H / 3:    PT(5).Y = Y + H / 2.1
            End If
            
        Case Style2
            GdipCreateSolidFill ConvertColor(vbBlack, 100), hBrush
            ReDim PT(5)
            If Top Then
                PT(0).X = X + H / 3:    PT(0).Y = Y + H / 1.66
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 2
                PT(2).X = X + H / 1.5:  PT(2).Y = Y + H / 1.66
                PT(3).X = X + H / 1.5:  PT(3).Y = Y + H / 1.7
                PT(4).X = X + H / 2:    PT(4).Y = Y + H / 3
                PT(5).X = X + H / 3:    PT(5).Y = Y + H / 1.7
            Else
                PT(0).X = X + H / 3:    PT(0).Y = Y + H / 2.5
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 2
                PT(2).X = X + H / 1.5:  PT(2).Y = Y + H / 2.5
                PT(3).X = X + H / 1.5:  PT(3).Y = Y + H / 2.4
                PT(4).X = X + H / 2:    PT(4).Y = Y + H / 1.5
                PT(5).X = X + H / 3:    PT(5).Y = Y + H / 2.4
            End If
            
        Case Style3
            GdipCreateSolidFill ConvertColor(vbBlack, 100), hBrush
            ReDim PT(2)
            If Top Then
                PT(0).X = X + H / 4:    PT(0).Y = Y + H / 1.5
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 3
                PT(2).X = X + H / 1.33: PT(2).Y = Y + H / 1.5
            Else
                PT(0).X = X + H / 4:    PT(0).Y = Y + H / 3
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 1.66
                PT(2).X = X + H / 1.33: PT(2).Y = Y + H / 3
            End If
            
        Case Style4
            GdipCreatePen1 ConvertColor(vbBlack, 100), 1, &H2, hPen
            ReDim PT(2)
            If Top Then
                PT(0).X = X + H / 4:    PT(0).Y = Y + H / 1.75
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 2.5
                PT(2).X = X + H / 1.33: PT(2).Y = Y + H / 1.75
            Else
                PT(0).X = X + H / 4:    PT(0).Y = Y + H / 2.33
                PT(1).X = X + H / 2:    PT(1).Y = Y + H / 1.66
                PT(2).X = X + H / 1.33: PT(2).Y = Y + H / 2.33
            End If
            
    End Select
    
    If m_eOrientation = oHorizontal Then
        Dim i As Long, T As Single
        For i = 0 To UBound(PT)
            T = PT(i).X
            PT(i).X = PT(i).Y
            PT(i).Y = T
        Next
    End If
    
    If hPen Then
        GdipDrawLines hGraphics, hPen, PT(0), UBound(PT) + 1
        GdipDeletePen hPen
    End If
    
    If hBrush Then
        GdipFillPolygon hGraphics, hBrush, PT(0), UBound(PT) + 1, FillModeAlternate
        GdipDeleteBrush hBrush
    End If
End Sub

Private Function IsMouseInControl() As Boolean
    Dim PT As POINTAPI

    Call GetCursorPos(PT)
    IsMouseInControl = CBool(WindowFromPoint(PT.X, PT.Y) = UserControl.hwnd)
End Function

Private Function IsMouseInContainer() As Boolean
    Dim PT As POINTAPI
    Dim REC As RECT
    GetWindowRect UserControl.ContainerHwnd, REC
    Call GetCursorPos(PT)
    IsMouseInContainer = CBool(PtInRect(REC, PT.X, PT.Y))
End Function

Public Function GetWindowsDPI() As Double
    Dim Hdc As Long, LPX  As Double
    Hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(Hdc, LOGPIXELSX))
    ReleaseDC 0, Hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim gdiSI           As GdiplusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    gdiSI.GdiplusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, gdiSI
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function


Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
End Function

