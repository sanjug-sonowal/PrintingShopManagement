VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type tagTHUMBBUTTON
    dwMask As THUMBBUTTONMASK
    iId As Long
    iBitmap As Long
    hIcon As Long
    szTip(259) As Integer
    dwFlags  As Long
End Type

Private Enum EnuTaskbarList3
    eSetProgressValue = 9
    eSetProgressState = 10
    eRegisterTab = 11
    eUnregisterTab = 12
    eSetTabOrder = 13
    eSetTabActive = 14
    eThumbBarAddButtons = 15
    eThumbBarUpdateButtons = 16
    eThumbBarSetImageList = 17
    eSetOverlayIcon = 18
    eSetThumbnailTooltip = 19
    eSetThumbnailClip = 20
End Enum

Private Enum THUMBBUTTONMASK
    THB_BITMAP = 1
    THB_ICON = 2
    THB_TOOLTIP = 4
    THB_FLAGS = 8
End Enum

Public Enum TBPFLAG
    TBPF_NOPROGRESS = 0
    TBPF_INDETERMINATE = 1
    TBPF_NORMAL = 2
    TBPF_ERROR = 4
    TBPF_PAUSED = 8
End Enum

Public Enum THUMBBUTTONFLAGS
    THBF_ENABLED = 0
    THBF_DISABLED = 1
    THBF_DISMISSONCLICK = 2
    THBF_NOBACKGROUND = 4
    THBF_HIDDEN = 8
    THBF_NONINTERACTIVE = &H10
End Enum

'Private Enum TBATFLAG
'    TBATF_USEMDITHUMBNAIL = 1
'    TBATF_USEMDILIVEPREVIEW = 2
'End Enum

Private Const CLSID_TaskbarList As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
'Private Const IID_ITaskbarList2 As String = "{602D4995-B13A-429b-A66E-1935E44F4317}"
'Private Const IID_ITaskbarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Const CLSID_TaskbarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
'Private Const CLSID_TaskbarList4 As String = "{C43DC798-95D1-4BEA-9030-BB99E2983A1A}"

Private Const IID_Release As Long = &H8&
Private Const CLSCTX_INPROC_SERVER      As Long = 1
Private Const CC_STDCALL As Long = &H4&
Private Const S_OK As Long = &H0&
Private iTaskBarVB3 As Long
Private Buttons() As tagTHUMBBUTTON
Private ButtonsCount As Long

Public Function SetProgressValue(hWnd As Long, ByVal Value As Currency, ByVal Max As Currency) As Boolean
    If iTaskBarVB3 Then
        SetProgressValue = DispCallByVtbl(iTaskBarVB3, eSetProgressValue, hWnd, Value, Max)
    End If
End Function

Public Function SetProgressState(hWnd As Long, tbpFlags As TBPFLAG) As Boolean
    If iTaskBarVB3 Then
        SetProgressState = DispCallByVtbl(iTaskBarVB3, eSetProgressState, hWnd, tbpFlags)
    End If
End Function

Public Function CreateButton(iId As Long, Optional hIcon As Long, Optional ByVal sTip As String, Optional Flags As THUMBBUTTONFLAGS) As Boolean
    Dim i As Long
    If iTaskBarVB3 Then
        ReDim Preserve Buttons(ButtonsCount)
        With Buttons(ButtonsCount)
            .iId = iId
            .hIcon = hIcon
            For i = 1 To Len(sTip)
                .szTip(i - 1) = AscW(Mid(sTip, i, 1))
            Next i
            .dwFlags = Flags
            .dwMask = (THB_ICON And CBool(hIcon)) Or (THB_TOOLTIP And CBool(Len(sTip))) Or THB_FLAGS
        End With
        ButtonsCount = ButtonsCount + 1
    End If
End Function

Public Function ThumbBarAddButtons(hWnd As Long) As Boolean
    ThumbBarAddButtons = DispCallByVtbl(iTaskBarVB3, eThumbBarAddButtons, hWnd, ButtonsCount, VarPtr(Buttons(0))) = S_OK
End Function

Public Function ThumbBarUpdateButtons(hWnd As Long, iId As Long, Optional hIcon As Long, Optional ByVal sTip As String, Optional Flags As THUMBBUTTONFLAGS) As Boolean
    Dim TB As tagTHUMBBUTTON, i As Long
    If iTaskBarVB3 Then
        With TB
            .iId = iId
            .hIcon = hIcon
            For i = 1 To Len(sTip)
                .szTip(i - 1) = AscW(Mid(sTip, i, 1))
            Next i
            .dwFlags = Flags
            .dwMask = (THB_ICON And CBool(hIcon)) Or (THB_TOOLTIP And CBool(Len(sTip))) Or THB_FLAGS
        End With

        ThumbBarUpdateButtons = DispCallByVtbl(iTaskBarVB3, eThumbBarUpdateButtons, hWnd, 1, VarPtr(TB)) = S_OK
    End If
End Function


Public Function SetThumbnailTooltip(hWnd As Long, pszTip As String) As Boolean
    If iTaskBarVB3 Then
        SetThumbnailTooltip = DispCallByVtbl(iTaskBarVB3, eSetThumbnailTooltip, hWnd, StrPtr(pszTip)) = S_OK
    End If
End Function

Public Function ResetThumbnailClip(hWnd As Long) As Boolean
      ResetThumbnailClip = DispCallByVtbl(iTaskBarVB3, eSetThumbnailClip, hWnd, 0&) = S_OK
End Function

Public Function SetThumbnailClip(hWnd As Long, Left As Long, Top As Long, Width As Long, Height As Long)
    Dim tRect As RECT
    
    If iTaskBarVB3 Then
        With tRect
            .Left = Left
            .Top = Top
            .Right = Left + Width
            .Bottom = Top + Height
        End With
        SetThumbnailClip = DispCallByVtbl(iTaskBarVB3, eSetThumbnailClip, hWnd, VarPtr(tRect)) = S_OK
    End If
End Function

Private Function DispCallByVtbl(pUnk As Long, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(pUnk, lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        err.Raise hResult, "DispCallFunc"
    End If
End Function

Public Function SetOverlayIcon(hWnd As Long, hIcon As Long, ByVal pszDescription As String) As Boolean
    If iTaskBarVB3 Then
        SetOverlayIcon = DispCallByVtbl(iTaskBarVB3, eSetOverlayIcon, hWnd, hIcon, VarPtr(pszDescription)) = S_OK
    End If
End Function

Private Sub ReleaseInterface(ByRef pInterface As Long)
    Dim vRet As Variant
    If pInterface <> 0& Then
        If DispCallFunc(pInterface, IID_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, vRet) = S_OK Then
            pInterface = 0&
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Dim iId As GUID
    Dim CLSID As GUID
    
    Call CLSIDFromString(StrPtr(CLSID_TaskbarList), CLSID)
    Call CLSIDFromString(StrPtr(CLSID_TaskbarList3), iId)
    Call CoCreateInstance(VarPtr(CLSID), 0&, CLSCTX_INPROC_SERVER, VarPtr(iId), iTaskBarVB3)
End Sub

Private Sub Class_Terminate()
    If iTaskBarVB3 Then ReleaseInterface iTaskBarVB3
End Sub
