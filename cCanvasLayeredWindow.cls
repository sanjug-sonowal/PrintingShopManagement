VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCanvasLayeredWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Type POINTAPI
   X                        As Long
   Y                        As Long
End Type

Private Type Size
   cx                       As Long
   cy                       As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAPINFOHEADER
   biSize                   As Long
   biWidth                  As Long
   biHeight                 As Long
   biPlanes                 As Integer
   biBitCount               As Integer
   biCompression            As Long
   biSizeImage              As Long
   biXPelsPerMeter          As Long
   biYPelsPerMeter          As Long
   biClrUsed                As Long
   biClrImportant           As Long
End Type

Private Type BITMAPINFO
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(3)             As Byte
End Type

Private Type BLENDFUNCTION
   BlendOp                  As Byte
   BlendFlags               As Byte
   SourceConstantAlpha      As Byte
   AlphaFormat              As Byte
End Type

Private Const ULW_ALPHA         As Long = &H2
Private Const BI_RGB            As Long = 0&
Private Const DIB_RGB_COLORS    As Long = 0&
Private Const AC_SRC_ALPHA      As Long = &H1

Private Const GWL_EXSTYLE       As Long = -20
Private Const WS_EX_TOPMOST     As Long = &H8&
Private Const WS_EX_LAYERED     As Long = &H80000

Private c_lhDC          As Long
Private c_lDIB          As Long
Private m_Pointer       As Long
Private m_OldBmp        As Long
Private tSIZE           As Size

Private tBLENDFUNCTION  As BLENDFUNCTION

Public Sub DestroyCanvas()
    If c_lDIB Then
        Call DeleteObject(SelectObject(c_lhDC, m_OldBmp)): c_lDIB = 0
        Call DeleteDC(c_lhDC): c_lhDC = 0
    End If
End Sub

Public Function CreateCanvas(ByVal Width As Long, ByVal Height As Long) As Boolean
    Dim tBITMAPINFO     As BITMAPINFO

    If c_lDIB <> 0 Then DestroyCanvas

    tSIZE.cx = Width
    tSIZE.cy = Height

    With tBITMAPINFO.bmiHeader
        .biSize = Len(tBITMAPINFO.bmiHeader)
        .biBitCount = 32
        .biHeight = tSIZE.cy
        .biWidth = tSIZE.cx
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * 4
    End With

    c_lhDC = CreateCompatibleDC(0)
    c_lDIB = CreateDIBSection(c_lhDC, tBITMAPINFO, DIB_RGB_COLORS, m_Pointer, 0&, 0&)
    m_OldBmp = SelectObject(c_lhDC, c_lDIB)

    CreateCanvas = c_lDIB <> 0
End Function

Public Sub Clear()
    If m_Pointer Then FillMemory ByVal m_Pointer, tSIZE.cx * tSIZE.cy * 4&, 0
End Sub

Public Property Get hdc() As Long
    hdc = c_lhDC
End Property

Public Sub UpdateLayered(hwnd As Long, Optional ByVal Alpha As Byte = 255)
    Dim tPT As POINTAPI
    Dim ExStyle As Long

    ExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (ExStyle And WS_EX_LAYERED) <> WS_EX_LAYERED Then
        Call SetWindowLong(hwnd, GWL_EXSTYLE, ExStyle Or WS_EX_LAYERED)
    End If
    
    With tBLENDFUNCTION
        .AlphaFormat = AC_SRC_ALPHA
        .SourceConstantAlpha = Alpha
    End With

    Call UpdateLayeredWindow(hwnd, 0&, ByVal 0&, tSIZE, c_lhDC, tPT, 0&, tBLENDFUNCTION, ULW_ALPHA)
End Sub

Private Sub Class_Terminate()
    DestroyCanvas
End Sub

