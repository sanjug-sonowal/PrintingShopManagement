Attribute VB_Name = "MdlImage"
Option Explicit
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mHbmReturn As Long, ByVal mBackground As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mHbmReturn As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long

Private Const CRYPT_STRING_BASE64           As Long = &H1

Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const UNIT_PIXELS = 2&
  
Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Private Type PICTDESC
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Public Function LoadPictureEx(SrcImg As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal bStretch As Boolean, Optional ByVal ReturnPicType As Long = vbPicTypeBitmap, Optional ByVal ForeColor As OLE_COLOR = -1, Optional ByVal BackColor As Long = -1) As IPicture
    Dim hImage1 As Long, hImage2 As Long, hGraphics As Long, hBitmap As Long
    Dim DataArr() As Byte
    Dim lPictureRealWidth As Long
    Dim lPictureRealHeight As Long
    Dim GdipToken As Long
    Dim x As Long, y As Long, cx As Long, cy As Long
    Dim sngRatio1 As Single, sngRatio2 As Single
    Dim tPicDesc As PICTDESC, GUID(0 To 3) As Long
    Dim hAttributes As Long
    Dim tMatrixColor    As COLORMATRIX, tMatrixGray    As COLORMATRIX
    
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)

    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    Call LoadImageFromArray(DataArr, hImage1)
                Else
                    Dim oXMLHTTP As Object
                    Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
                    
                    oXMLHTTP.Open "GET", SrcImg, True
                    oXMLHTTP.send
                    While oXMLHTTP.readyState <> 4
                        DoEvents
                    Wend
                    If oXMLHTTP.Status = 200 Then
                        DataArr() = oXMLHTTP.responseBody
                        Call LoadImageFromArray(DataArr, hImage1)
                    End If
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), hImage1)
            End If
        Case vbLong
            Const OBJ_BITMAP As Long = 7
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, hImage1)
            Else
                Call GdipCreateBitmapFromHICON(SrcImg, hImage1)
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, hImage1)
            
        Case (vbArray Or vbByte)
            DataArr() = SrcImg
            Call LoadImageFromArray(DataArr, hImage1)
    End Select

    If hImage1 <> 0 Then
        GdipGetImageWidth hImage1, lPictureRealWidth
        GdipGetImageHeight hImage1, lPictureRealHeight
        If Width = 0 Then Width = lPictureRealWidth
        If Height = 0 Then Height = lPictureRealWidth
        
        If bStretch = False Then
            sngRatio1 = Width / lPictureRealWidth
            sngRatio2 = Height / lPictureRealHeight
            If sngRatio1 > sngRatio2 Then sngRatio1 = sngRatio2
            cx = lPictureRealWidth * sngRatio1: cy = lPictureRealHeight * sngRatio1
            x = (Width - cx) \ 2: y = (Height - cy) \ 2
        Else
            cx = Width: cy = Height: x = 0: y = 0
        End If
        
        GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        
        
        If ForeColor <> -1 Then
            Dim R As Byte, G As Byte, B As Byte
                
            Call GdipCreateImageAttributes(hAttributes)
            If (ForeColor And &H80000000) Then ForeColor = GetSysColor(ForeColor And &HFF&)
            
            With tMatrixColor
                B = ((ForeColor \ &H10000) And &HFF)
                G = ((ForeColor \ &H100) And &HFF)
                R = (ForeColor And &HFF)
                
                .m(0, 0) = R / 255
                .m(1, 0) = G / 255
                .m(2, 0) = B / 255
                .m(0, 4) = R / 255
                .m(1, 4) = G / 255
                .m(2, 4) = B / 255
                .m(3, 3) = 1 'm_PictureOpacity / 100
                .m(4, 4) = 1
            End With
        
            GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
        End If
        
        GdipDrawImageRectRectI hGraphics, hImage1, x, y, cx, cy, 0, 0, lPictureRealWidth, lPictureRealHeight, UNIT_PIXELS, hAttributes, 0, 0
        If hAttributes <> 0 Then GdipDisposeImageAttributes hAttributes
        GdipDeleteGraphics hGraphics
        GdipDisposeImage hImage1

        If ReturnPicType = vbPicTypeBitmap Then
           If BackColor = -1 Then BackColor = 0 Else BackColor = RGBtoARGB(BackColor, 100)
            GdipCreateHBITMAPFromBitmap hImage2, hBitmap, BackColor
        Else
            GdipCreateHICONFromBitmap hImage2, hBitmap
        End If
        GdipDisposeImage hImage2
        
        ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        GUID(0) = &H7BF80980: GUID(1) = &H101ABF32: GUID(2) = &HAA00BB8B: GUID(3) = &HAB0C3000
        
        With tPicDesc
            .Size = Len(tPicDesc)
            .type = ReturnPicType
            .hBmp = hBitmap
        End With

        Call OleCreatePictureIndirect(tPicDesc, GUID(0), True, LoadPictureEx)
    End If
    
    Call GdiplusShutdown(GdipToken)
        
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
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


Private Function LoadImageFromArray(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromArray_Error
    
    Dim IStream     As stdole.IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromArray = True
        End If
    End If

    Set IStream = Nothing
    
LoadImageFromArray_Error:
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function


Private Function Base64Decode(ByVal sIn As String, ByRef bvOut() As Byte) As Boolean
                              
    Dim lLenOut                 As Long
    '// calculate buffer len
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, 0, VarPtr(lLenOut), 0, 0)
 
    If lLenOut = 0 Then
        Exit Function
    End If
 
    ReDim bvOut(lLenOut - 1)
    '// now convert to base64
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, VarPtr(bvOut(0)), VarPtr(lLenOut), 0, 0)
    Base64Decode = True
End Function

Public Function GetWindowsDPI() As Double
    Const LOGPIXELSX = 88&
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
