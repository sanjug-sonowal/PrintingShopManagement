VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsShellItemProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'DATE: 05/11/2019
'I: LEANDRO ASCIERTO www.leandroascierto.com
'KEY REFERERCES: https://github.com/aybe/Windows-API-Code-Pack-1.1/blob/master/source/WindowsAPICodePack/Shell/PropertySystem/SystemProperties.cs
'REQUERIMENTES: -Windows Vista or newer -oleexp.tlb 4.6 or higher http://www.vbforums.com/showthread.php?786079-VB6-Modern-Shell-Interface-Type-Library-oleexp-tlb
'MADE BASE: on the initial "FAFALONE" project: http://www.vbforums.com/showthread.php?880337-VB6-Write-MP3-Album-Art-and-other-tags-using-the-Windows-Property-System
'NOTES:
'MANY OF THE PROPERTIES WERE NOT TESTED, EVEN MANY ARE READING ONLY
'THE PROPERTY m_ThumbnailData IT WAS CREATED BECAUSE m_ThumbnailStream COULD NOT BE SAVED IF IT WAS READ BY GDI+

Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As Any) As Long
'Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function VariantToPropVariant Lib "propsys.dll" (ByRef pvar As Variant, ByRef ppropvar As Variant) As Long
Private Declare Function PropVariantToVariant Lib "propsys" (ByRef propvar As Any, ByRef Var As Variant) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function SHCreateItemFromParsingName Lib "Shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, riid As GUID, pUnk As Long) As Long
Private Declare Function SHGetPropertyStoreFromParsingName Lib "Shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, ByVal Flags As GETPROPERTYSTOREFLAGS, ByRef riid As GUID, ByRef ppv As Long) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PropertyKey
    fmtid As GUID
    pid As Long
End Type

Private Enum GETPROPERTYSTOREFLAGS
    GPS_DEFAULT = &H0
    GPS_HANDLERPROPERTIESONLY = &H1
    GPS_READWRITE = &H2
    GPS_TEMPORARY = &H4
    GPS_FASTPROPERTIESONLY = &H8
    GPS_OPENSLOWITEM = &H10
    GPS_DELAYCREATION = &H20
    GPS_BESTEFFORT = &H40
    GPS_NO_OPLOCK = &H80
    GPS_MASK_VALID = &HFF
End Enum

Private Const PKEY_Audio_ChannelCount = "{64440490-4C8B-11D1-8B70-080036B11A03},7 "
Private Const PKEY_Audio_EncodingBitrate = "{64440490-4C8B-11D1-8B70-080036B11A03},4 "
Private Const PKEY_Audio_PeakValue = "{2579E5D0-1116-4084-BD9A-9B4F7CB4DF5E},100  "
Private Const PKEY_Audio_SampleRate = "{64440490-4C8B-11D1-8B70-080036B11A03},5 "
Private Const PKEY_Author = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},4 "
Private Const PKEY_Comment = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},6 "
Private Const PKEY_Copyright = "{64440492-4C8B-11D1-8B70-080036B11A03},11"
Private Const PKEY_Keywords = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},5 "
Private Const PKEY_Language = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE},28"
Private Const PKEY_Media_AverageLevel = "{09EDD5B6-B301-43C5-9990-D00302EFFD46},100  "
Private Const PKEY_Media_DateEncoded = "{2E4B640D-5019-46D8-8881-55414CC5CAA0},100  "
Private Const PKEY_Media_Duration = "{64440490-4C8B-11D1-8B70-080036B11A03},3 "
Private Const PKEY_Media_EncodedBy = "{64440492-4C8B-11D1-8B70-080036B11A03},36"
Private Const PKEY_Media_ProviderStyle = "{64440492-4C8B-11D1-8B70-080036B11A03},40"
Private Const PKEY_Media_Publisher = "{64440492-4C8B-11D1-8B70-080036B11A03},30"
Private Const PKEY_Media_SubTitle = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},38"
Private Const PKEY_Media_Writer = "{64440492-4C8B-11D1-8B70-080036B11A03}, 23  "
Private Const PKEY_Media_Year = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},5 "
Private Const PKEY_Music_AlbumArtist = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},13"
Private Const PKEY_Music_AlbumTitle = "{56A3372E-CE9C-11D2-9F0E-006097C686F6}, 4"
Private Const PKEY_Music_Artist = "{56A3372E-CE9C-11D2-9F0E-006097C686F6}, 2"
Private Const PKEY_Music_Composer = "{64440492-4C8B-11D1-8B70-080036B11A03}, 19  "
Private Const PKEY_Music_Conductor = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},36"
Private Const PKEY_Music_Genre = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},11"
Private Const PKEY_Music_Lyrics = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},12"
Private Const PKEY_Music_Mood = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},39"
Private Const PKEY_Music_TrackNumber = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},7 "
Private Const PKEY_ParentalRating = " {10984E0A-F9F2-4321-B7EF-BAF195AF4319}, 100 "
Private Const PKEY_Rating = "{64440492-4C8B-11D1-8B70-080036B11A03},9 "
Private Const PKEY_ThumbnailStream = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
Private Const PKEY_Title = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},2 "

Public Enum ePropertyKey
    Audio_ChannelCount
    Audio_EncodingBitrate
    Audio_PeakValue
    Audio_SampleRate
    Author
    Comment
    Copyright
    Keywords
    Language
    Media_AverageLevel
    Media_DateEncoded
    Media_Duration
    Media_EncodedBy
    Media_ProviderStyle
    Media_Publisher
    Media_SubTitle
    Media_Writer
    Media_Year
    Music_AlbumArtist
    Music_AlbumTitle
    Music_Artist
    Music_Composer
    Music_Conductor
    Music_Genre
    Music_Lyrics
    Music_Mood
    Music_TrackNumber
    ParentalRating
    Rating
    ThumbnailStream
    Title
End Enum

Private Const VT_STREAM  As Long = &H42
Private Const CC_STDCALL As Long = &H4&
Private Const IID_Release As Long = &H8&
Private Const S_OK As Long = &H0&

Private Const ShellItem2 = "{7e9fb0d3-919f-4307-ab2e-9b1860310c93}"
Private Const PropertyStore = "{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}"

Private Const ifGetProperty = 13
Private Const ifCommit = 7
Private Const ifSetValue = 6

Private IID_IShellItem2 As GUID
Private IID_IPropertyStore As GUID

Private Sub Class_Initialize()
    CLSIDFromString StrPtr(ShellItem2), IID_IShellItem2
    CLSIDFromString StrPtr(PropertyStore), IID_IPropertyStore
End Sub

Public Function GetFileProperty(sPath As String, PropertyKey As ePropertyKey, Optional StrSeparators As String) As Variant
    Dim lPtrShellItem As Long 'IShellItem2
    Dim tPK As PropertyKey
    Dim vProp As Variant, vResult As Variant
        
    tPK = PropKeyFromString(GetPropKeyString(PropertyKey))
    
    Call SHCreateItemFromParsingName(StrPtr(sPath), 0&, IID_IShellItem2, lPtrShellItem)
    
    If lPtrShellItem Then
        Call DispCallByVtbl(lPtrShellItem, ifGetProperty, VarPtr(tPK), VarPtr(vProp))
        PropVariantToVariant vProp, vResult
        If Len(StrSeparators) Then
            If VarType(vResult) = (vbArray Or vbString) Then
                GetFileProperty = Join(vResult, StrSeparators)
            Else
                GetFileProperty = vResult
            End If
        Else
            GetFileProperty = vResult
        End If
        ReleaseInterface lPtrShellItem
    End If
End Function

Public Function SetFileProperty(sPath As String, PropertyKey As ePropertyKey, NewValue As Variant) As Boolean
    Dim tPK As PropertyKey
    Dim lPtrPropertyStore As Long 'IPropertyStore
    
    tPK = PropKeyFromString(GetPropKeyString(PropertyKey))
    
    Call SHGetPropertyStoreFromParsingName(StrPtr(sPath), 0&, GPS_HANDLERPROPERTIESONLY Or GPS_READWRITE, IID_IPropertyStore, lPtrPropertyStore)

    If lPtrPropertyStore Then
        If DispCallByVtbl(lPtrPropertyStore, ifSetValue, VarPtr(tPK), VarPtr(NewValue)) = S_OK Then
            SetFileProperty = DispCallByVtbl(lPtrPropertyStore, ifCommit) = S_OK
        End If
        ReleaseInterface lPtrPropertyStore
    End If
End Function

Public Function SetFileThumbnail(sPath As String, bData() As Byte) As Boolean
    Dim vt As Integer
    Dim ThumbnailStream As Variant
    Dim lPtrPropertyStore As Long 'IPropertyStore
    
    Call SHGetPropertyStoreFromParsingName(StrPtr(sPath), 0&, GPS_HANDLERPROPERTIESONLY Or GPS_READWRITE, IID_IPropertyStore, lPtrPropertyStore)
    
    If lPtrPropertyStore Then
    
        StreamFromGlobal bData, ThumbnailStream
    
        If Not VarType(ThumbnailStream) = vbEmpty Then
            vt = VT_STREAM
            CopyMemory ByVal VarPtr(ThumbnailStream), ByVal VarPtr(vt), 2&
        End If
        
        If DispCallByVtbl(lPtrPropertyStore, ifSetValue, VarPtr(PropKeyFromString(PKEY_ThumbnailStream)), VarPtr(ThumbnailStream)) = S_OK Then
            SetFileThumbnail = DispCallByVtbl(lPtrPropertyStore, ifCommit) = S_OK
        End If
        ReleaseInterface lPtrPropertyStore
    End If
End Function

Public Function GetFileThumbnail(sPath As String) As IUnknown
    Dim lPtrShellItem As Long 'IShellItem2
    Dim IID_IShellItem2 As GUID
    Dim vProp As Variant, ThumbnailStream As Variant
        
    CLSIDFromString StrPtr(ShellItem2), IID_IShellItem2
    Call SHCreateItemFromParsingName(StrPtr(sPath), 0&, IID_IShellItem2, lPtrShellItem)
    
    If lPtrShellItem Then
        Call DispCallByVtbl(lPtrShellItem, ifGetProperty, VarPtr(PropKeyFromString(PKEY_ThumbnailStream)), VarPtr(vProp))
        If VarType(vProp) = VT_STREAM Then
            PropVariantToVariant vProp, ThumbnailStream
            If VarType(ThumbnailStream) = vbDataObject Then
                Set GetFileThumbnail = ThumbnailStream
            End If
        End If
        ReleaseInterface lPtrShellItem
        
    End If
End Function

Private Function GetPropKeyString(PropertyKey As ePropertyKey) As String
    Dim sRet As String
    Select Case PropertyKey
        Case Audio_ChannelCount: sRet = PKEY_Audio_ChannelCount
        Case Audio_EncodingBitrate: sRet = PKEY_Audio_EncodingBitrate
        Case Audio_PeakValue: sRet = PKEY_Audio_PeakValue
        Case Audio_SampleRate: sRet = PKEY_Audio_SampleRate
        Case Author: sRet = PKEY_Author
        Case Comment: sRet = PKEY_Comment
        Case Copyright: sRet = PKEY_Copyright
        Case Keywords: sRet = PKEY_Keywords
        Case Language: sRet = PKEY_Language
        Case Media_AverageLevel: sRet = PKEY_Media_AverageLevel
        Case Media_DateEncoded: sRet = PKEY_Media_DateEncoded
        Case Media_Duration: sRet = PKEY_Media_Duration
        Case Media_EncodedBy: sRet = PKEY_Media_EncodedBy
        Case Media_ProviderStyle: sRet = PKEY_Media_ProviderStyle
        Case Media_Publisher: sRet = PKEY_Media_Publisher
        Case Media_SubTitle: sRet = PKEY_Media_SubTitle
        Case Media_Writer: sRet = PKEY_Media_Writer
        Case Media_Year: sRet = PKEY_Media_Year
        Case Music_AlbumArtist: sRet = PKEY_Music_AlbumArtist
        Case Music_AlbumTitle: sRet = PKEY_Music_AlbumTitle
        Case Music_Artist: sRet = PKEY_Music_Artist
        Case Music_Composer: sRet = PKEY_Music_Composer
        Case Music_Conductor: sRet = PKEY_Music_Conductor
        Case Music_Genre: sRet = PKEY_Music_Genre
        Case Music_Lyrics: sRet = PKEY_Music_Lyrics
        Case Music_Mood: sRet = PKEY_Music_Mood
        Case Music_TrackNumber: sRet = PKEY_Music_TrackNumber
        Case ParentalRating: sRet = PKEY_ParentalRating
        Case Rating: sRet = PKEY_Rating
        Case ThumbnailStream: sRet = PKEY_ThumbnailStream
        Case Title: sRet = PKEY_Title
    End Select
    GetPropKeyString = sRet
End Function

Private Function PropKeyFromString(sValue As String) As PropertyKey
    Dim KeyID() As String
    KeyID = Split(sValue, ",")
    CLSIDFromString StrPtr(KeyID(0)), PropKeyFromString.fmtid
    PropKeyFromString.pid = KeyID(1)
End Function

Private Sub ReleaseInterface(ByRef pInterface As Long)
    Dim vRet As Variant
    If pInterface <> 0& Then
        If DispCallFunc(pInterface, IID_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, vRet) = S_OK Then
            pInterface = 0&
        End If
    End If
End Sub

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

Private Function StreamFromGlobal(Data() As Byte, IStream) As Boolean
    Dim vbr As Variant
    Dim oStrm As IUnknown
    
    If Not IsArrayDim(VarPtrArray(Data)) Then Exit Function

    Call CreateStreamOnHGlobal(Data(0), 1, oStrm)

    Set vbr = oStrm
    VariantToPropVariant vbr, IStream

End Function

'Private Function GlobalFromStream(IStream, Data() As Byte)
'    Dim hMem As Long, lSize As Long
'    Dim lPt As Long
'
'    If GetHGlobalFromStream(ObjPtr(IStream), hMem) = 0 Then
'        lSize = GlobalSize(hMem)
'        If lSize Then
'            lPt = GlobalLock(hMem)
'            ReDim Data(0 To lSize - 1)
'            CopyMemory Data(0), ByVal lPt, lSize
'            GlobalUnlock hMem
'        End If
'    End If
'End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function


