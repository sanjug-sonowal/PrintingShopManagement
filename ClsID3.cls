VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsID3"
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
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function VariantToPropVariant Lib "propsys.dll" (ByRef pvar As Variant, ByRef ppropvar As Variant) As Long
Private Declare Function PropVariantToVariant Lib "propsys" (ByRef propvar As Any, ByRef Var As Variant) As Long
    
Private Const FILE_READ_DATA = &H1
Private Const FILE_SHARE_READ = &H1&
Private Const OPEN_EXISTING = 3&


Private Const PKEY_ThumbnailStream = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},27"
Private Const PKEY_Title = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},2"
Private Const PKEY_Author = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},4"
Private Const PKEY_Comment = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},6"
Private Const PKEY_Copyright = "{64440492-4C8B-11D1-8B70-080036B11A03},11"
Private Const PKEY_Rating = "{64440492-4C8B-11D1-8B70-080036B11A03},9"
Private Const PKEY_Keywords = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9},5"
Private Const PKEY_Language = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE},28"
Private Const PKEY_ParentalRating = " {10984E0A-F9F2-4321-B7EF-BAF195AF4319}, 100"
Private Const PKEY_Audio_EncodingBitrate = "{64440490-4C8B-11D1-8B70-080036B11A03},4"
Private Const PKEY_Audio_ChannelCount = "{64440490-4C8B-11D1-8B70-080036B11A03},7"
Private Const PKEY_Audio_PeakValue = "{2579E5D0-1116-4084-BD9A-9B4F7CB4DF5E},100"
Private Const PKEY_Audio_SampleRate = "{64440490-4C8B-11D1-8B70-080036B11A03},5"
Private Const PKEY_Media_SubTitle = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},38"
Private Const PKEY_Media_ProviderStyle = "{64440492-4C8B-11D1-8B70-080036B11A03},40"
Private Const PKEY_Media_Year = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},5"
Private Const PKEY_Media_DateEncoded = "{2E4B640D-5019-46D8-8881-55414CC5CAA0},100"
Private Const PKEY_Media_Writer = "{64440492-4C8B-11D1-8B70-080036B11A03}, 23"
Private Const PKEY_Media_AverageLevel = "{09EDD5B6-B301-43C5-9990-D00302EFFD46},100"
Private Const PKEY_Media_Duration = "{64440490-4C8B-11D1-8B70-080036B11A03},3"
Private Const PKEY_Media_Publisher = "{64440492-4C8B-11D1-8B70-080036B11A03},30"
Private Const PKEY_Media_EncodedBy = "{64440492-4C8B-11D1-8B70-080036B11A03},36"
Private Const PKEY_Music_AlbumTitle = "{56A3372E-CE9C-11D2-9F0E-006097C686F6}, 4"
Private Const PKEY_Music_Artist = "{56A3372E-CE9C-11D2-9F0E-006097C686F6}, 2"
Private Const PKEY_Music_Composer = "{64440492-4C8B-11D1-8B70-080036B11A03}, 19"
Private Const PKEY_Music_Conductor = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},36"
Private Const PKEY_Music_Lyrics = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},12"
Private Const PKEY_Music_Mood = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},39"
Private Const PKEY_Music_TrackNumber = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},7"
Private Const PKEY_Music_Genre = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},11"
Private Const PKEY_Music_AlbumArtist = "{56A3372E-CE9C-11D2-9F0E-006097C686F6},13"

Private Const ShellItem2 = "{7e9fb0d3-919f-4307-ab2e-9b1860310c93}"
Private Const PropertyStoreFactory = "{bc110b6d-57e8-4148-a9c6-91015ab2f3a5}"
Private Const PropertyStore = "{886d8eeb-8cf2-4446-8d02-cdba1dbdcf99}"

Private m_ThumbnailData() As Byte
Private m_ThumbnailStream As Variant
Private m_Title As String
Private m_Author As Variant
Private m_Comment As String
Private m_Copyright As String
Private m_Rating As String
Private m_Keywords As String
Private m_Language As String
Private m_ParentalRating As String
Private m_EncodingBitrate As String
Private m_ChannelCount As String
Private m_PeakValue As String
Private m_SampleRate As String
Private m_SubTitle As String
Private m_ProviderStyle As String
Private m_Year As Variant
Private m_DateEncoded As String
Private m_Writer As String
Private m_AverageLevel As String
Private m_Duration As String
Private m_Publisher As String
Private m_EncodedBy As String
Private m_AlbumTitle As String
Private m_Artist As Variant
Private m_Composer As Variant
Private m_Conductor As String
Private m_Lyrics As String
Private m_Mood As String
Private m_TrackNumber As String
Private m_Genre As Variant
Private m_AlbumArtist As String
Private m_FilePath As String

Private Const ifGetProperty = 13
Private Const ifGetPropertyStore = 8
Private Const ifCommit = 7
Private Const ifSetValue = 6

Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function SHCreateItemFromParsingName Lib "Shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, riid As GUID, pUnk As Long) As Long
Private Declare Function SHGetPropertyStoreFromParsingName Lib "Shell32.dll" (ByVal pszPath As Long, ByVal pBC As Long, ByVal Flags As GETPROPERTYSTOREFLAGS, ByRef riid As GUID, ByRef ppv As Long) As Long
                        
Public Enum GETPROPERTYSTOREFLAGS
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

Private Const VT_STREAM  As Long = &H42
Private Const CC_STDCALL As Long = &H4&
Private Const IID_Release As Long = &H8&
Private Const S_OK As Long = &H0&
Dim lPtrShellItem As Long 'IShellItem2
Dim lPtrPropertyStore As Long 'IPropertyStore

Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property


Private Function PropKeyFromString(sValue As String) As PropertyKey
    Dim KeyID() As String
    KeyID = Split(sValue, ",")
    CLSIDFromString StrPtr(KeyID(0)), PropKeyFromString.fmtid
    PropKeyFromString.pid = KeyID(1)
End Function


Public Function GetFileImageStream(sPath As String) As IUnknown
    Dim IID_IShellItem2 As GUID
    Dim vProp As Variant
        
    CLSIDFromString StrPtr(ShellItem2), IID_IShellItem2
    Call SHCreateItemFromParsingName(StrPtr(sPath), 0&, IID_IShellItem2, lPtrShellItem)
    
    If lPtrShellItem Then
        
        Call DispCallByVtbl(lPtrShellItem, ifGetProperty, VarPtr(PropKeyFromString(PKEY_ThumbnailStream)), VarPtr(vProp))
        PropVariantToVariant vProp, m_ThumbnailStream
        Set GetFileImageStream = m_ThumbnailStream
        ReleaseInterface lPtrShellItem
    
    End If
End Function


Public Function LoadFile(ByVal sPath As String) As Boolean
    Dim IID_IShellItem2 As GUID
    Dim vProp As Variant
    
    m_FilePath = sPath
    
    CLSIDFromString StrPtr(ShellItem2), IID_IShellItem2
    Call SHCreateItemFromParsingName(StrPtr(sPath), 0&, IID_IShellItem2, lPtrShellItem)
    
    If lPtrShellItem Then
        
        Call DispCallByVtbl(lPtrShellItem, ifGetProperty, VarPtr(PropKeyFromString(PKEY_ThumbnailStream)), VarPtr(vProp))
        PropVariantToVariant vProp, m_ThumbnailStream

        If VarType(m_ThumbnailStream) = vbDataObject Then
            GlobalFromStream m_ThumbnailStream, m_ThumbnailData
        End If

        m_Title = GetProperty(PKEY_Title)
        m_Author = GetProperty(PKEY_Author)
        m_Comment = GetProperty(PKEY_Comment) 'READING FAILS, THIS CAN BE SEEN IN THE SAME WINDOWS EXPLORER
        m_Copyright = GetProperty(PKEY_Copyright)
        m_Rating = GetProperty(PKEY_Rating)
        m_Keywords = GetProperty(PKEY_Keywords)
        m_Language = GetProperty(PKEY_Language)
        m_ParentalRating = GetProperty(PKEY_ParentalRating)
        m_EncodingBitrate = GetProperty(PKEY_Audio_EncodingBitrate)
        m_ChannelCount = GetProperty(PKEY_Audio_ChannelCount)
        m_PeakValue = GetProperty(PKEY_Audio_PeakValue)
        m_SampleRate = GetProperty(PKEY_Audio_SampleRate)
        m_SubTitle = GetProperty(PKEY_Media_SubTitle)
        m_ProviderStyle = GetProperty(PKEY_Media_ProviderStyle)
        m_Year = GetProperty(PKEY_Media_Year)
        m_DateEncoded = GetProperty(PKEY_Media_DateEncoded)
        m_Writer = GetProperty(PKEY_Media_Writer)
        m_AverageLevel = GetProperty(PKEY_Media_AverageLevel)
        m_Duration = GetProperty(PKEY_Media_Duration)
        m_Publisher = GetProperty(PKEY_Media_Publisher)
        m_EncodedBy = GetProperty(PKEY_Media_EncodedBy)
        m_AlbumTitle = GetProperty(PKEY_Music_AlbumTitle)
        m_Artist = GetProperty(PKEY_Music_Artist)
        m_Composer = GetProperty(PKEY_Music_Composer)
        m_Conductor = GetProperty(PKEY_Music_Conductor)
        m_Lyrics = GetProperty(PKEY_Music_Lyrics)
        m_Mood = GetProperty(PKEY_Music_Mood)
        m_TrackNumber = GetProperty(PKEY_Music_TrackNumber)
        m_Genre = GetProperty(PKEY_Music_Genre)
        m_AlbumArtist = GetProperty(PKEY_Music_AlbumArtist)
        LoadFile = True
        
        ReleaseInterface lPtrShellItem
    End If
    
End Function


Public Function WriteFile() As Boolean
    Dim IID_IPropertyStore As GUID, vt As Integer
    
    
    CLSIDFromString StrPtr(PropertyStore), IID_IPropertyStore
    Call SHGetPropertyStoreFromParsingName(StrPtr(m_FilePath), 0&, GPS_HANDLERPROPERTIESONLY Or GPS_READWRITE, IID_IPropertyStore, lPtrPropertyStore)
    'If DispCallByVtbl(lPtrShellItem, ifGetPropertyStore, GPS_OPENSLOWITEM Or GPS_READWRITE, VarPtr(IID_IPropertyStore), VarPtr(lPtrPropertyStore)) Then

    
    If lPtrPropertyStore Then
        
        'FOR ANY REASON, WHEN THE ARTIST PROPERTY IS EMPTY, IT MUST BE WRITTEN AND SAVED TWICE!!
        Call SetValue(PKEY_Music_Artist, m_Artist)
        Call DispCallByVtbl(lPtrPropertyStore, ifCommit)
       

        StreamFromGlobal m_ThumbnailData, m_ThumbnailStream
    
        If Not VarType(m_ThumbnailStream) = vbEmpty Then
            vt = VT_STREAM
            CopyMemory ByVal VarPtr(m_ThumbnailStream), ByVal VarPtr(vt), 2&
        End If

        SetValue PKEY_ThumbnailStream, m_ThumbnailStream
        SetValue PKEY_Title, m_Title
        SetValue PKEY_Author, m_Author
        SetValue PKEY_Comment, m_Comment
        SetValue PKEY_Copyright, m_Copyright
        SetValue PKEY_Rating, m_Rating
        SetValue PKEY_Keywords, m_Keywords
        SetValue PKEY_Language, m_Language
        SetValue PKEY_ParentalRating, m_ParentalRating
        SetValue PKEY_Audio_EncodingBitrate, m_EncodingBitrate
        SetValue PKEY_Audio_ChannelCount, m_ChannelCount
        SetValue PKEY_Audio_PeakValue, m_PeakValue
        SetValue PKEY_Audio_SampleRate, m_SampleRate
        SetValue PKEY_Media_SubTitle, m_SubTitle
        SetValue PKEY_Media_ProviderStyle, m_ProviderStyle
        SetValue PKEY_Media_Year, m_Year
        SetValue PKEY_Media_DateEncoded, m_DateEncoded
        SetValue PKEY_Media_Duration, m_Duration
        SetValue PKEY_Media_Publisher, m_Publisher
        SetValue PKEY_Media_EncodedBy, m_EncodedBy
        SetValue PKEY_Music_AlbumTitle, m_AlbumTitle
        SetValue PKEY_Music_Artist, m_Artist
        SetValue PKEY_Music_Composer, m_Composer
        SetValue PKEY_Music_Conductor, m_Conductor
        SetValue PKEY_Music_Lyrics, m_Lyrics
        SetValue PKEY_Music_Mood, m_Mood
        SetValue PKEY_Music_TrackNumber, m_TrackNumber
        SetValue PKEY_Music_Genre, m_Genre
        SetValue PKEY_Music_AlbumArtist, m_AlbumArtist

        WriteFile = DispCallByVtbl(lPtrPropertyStore, ifCommit) = S_OK
        ReleaseInterface lPtrPropertyStore
    End If
End Function

Private Function SetValue(sPropKey As String, vValue As Variant) As Boolean
 SetValue = DispCallByVtbl(lPtrPropertyStore, ifSetValue, VarPtr(PropKeyFromString(sPropKey)), VarPtr(vValue)) = 0
End Function

Private Function GetProperty(sPropKey As String) As Variant
    Dim vProp As Variant
    Call DispCallByVtbl(lPtrShellItem, ifGetProperty, VarPtr(PropKeyFromString(sPropKey)), VarPtr(vProp))
    PropVariantToVariant vProp, GetProperty
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
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Private Function StreamFromGlobal(Data() As Byte, IStream) As Boolean
    Dim vbr As Variant, vpr As Variant
    Dim oStrm As IUnknown
    
    If Not IsArrayDim(VarPtrArray(Data)) Then Exit Function

    Call CreateStreamOnHGlobal(Data(0), 1, oStrm)

    Set vbr = oStrm
    VariantToPropVariant vbr, IStream

End Function

Private Function GlobalFromStream(IStream, Data() As Byte)
    Dim hMem As Long, lSize As Long
    Dim lPt As Long
    
    If GetHGlobalFromStream(ObjPtr(IStream), hMem) = 0 Then
        lSize = GlobalSize(hMem)
        If lSize Then
            lPt = GlobalLock(hMem)
            ReDim Data(0 To lSize - 1)
            CopyMemory Data(0), ByVal lPt, lSize
            GlobalUnlock hMem
        End If
    End If
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Public Property Get ThumbnailData() As Byte():   Let ThumbnailData = m_ThumbnailData: End Property
Public Property Get ThumbnailStream() As Variant:  ThumbnailStream = m_ThumbnailStream: End Property
Public Property Get Title() As String: Title = m_Title: End Property
Public Property Get Author() As Variant: Author = m_Author: End Property
Public Property Get Comment() As String: Comment = m_Comment: End Property
Public Property Get Copyright() As String: Copyright = m_Copyright: End Property
Public Property Get Rating() As String: Rating = m_Rating: End Property
Public Property Get Keywords() As String: Keywords = m_Keywords: End Property
Public Property Get Language() As String: Language = m_Language: End Property
Public Property Get ParentalRating() As String: ParentalRating = m_ParentalRating: End Property
Public Property Get EncodingBitrate() As String: EncodingBitrate = m_EncodingBitrate: End Property
Public Property Get ChannelCount() As String: ChannelCount = m_ChannelCount: End Property
Public Property Get PeakValue() As String: PeakValue = m_PeakValue: End Property
Public Property Get SampleRate() As String: SampleRate = m_SampleRate: End Property
Public Property Get SubTitle() As String: SubTitle = m_SubTitle: End Property
Public Property Get ProviderStyle() As String: ProviderStyle = m_ProviderStyle: End Property
Public Property Get Year() As Variant: Year = m_Year: End Property
Public Property Get DateEncoded() As String: DateEncoded = m_DateEncoded: End Property
Public Property Get Writer() As String: Writer = m_Writer: End Property
Public Property Get AverageLevel() As String: AverageLevel = m_AverageLevel: End Property
Public Property Get Duration() As String: Duration = m_Duration: End Property
Public Property Get Publisher() As String: Publisher = m_Publisher: End Property
Public Property Get EncodedBy() As String: EncodedBy = m_EncodedBy: End Property
Public Property Get AlbumTitle() As String: AlbumTitle = m_AlbumTitle: End Property
Public Property Get Artist() As Variant: Artist = m_Artist: End Property
Public Property Get Composer() As Variant: Composer = m_Composer: End Property
Public Property Get Conductor() As String: Conductor = m_Conductor: End Property
Public Property Get Lyrics() As String: Lyrics = m_Lyrics: End Property
Public Property Get Mood() As String: Mood = m_Mood: End Property
Public Property Get TrackNumber() As String: TrackNumber = m_TrackNumber: End Property
Public Property Get Genre() As Variant: Genre = m_Genre: End Property
Public Property Get AlbumArtist() As String: AlbumArtist = m_AlbumArtist: End Property

Public Property Let ThumbnailData(Value() As Byte):   Let m_ThumbnailData = Value: End Property
Public Property Let ThumbnailStream(ByVal Value As Variant):  m_ThumbnailStream = Value: End Property
Public Property Let Title(ByVal Value As String): m_Title = Value: End Property
Public Property Let Author(ByVal Value As Variant): m_Author = Value: End Property
Public Property Let Comment(ByVal Value As String): m_Comment = Value: End Property
Public Property Let Copyright(ByVal Value As String): m_Copyright = Value: End Property
Public Property Let Rating(ByVal Value As String): m_Rating = Value: End Property
Public Property Let Keywords(ByVal Value As String): m_Keywords = Value: End Property
Public Property Let Language(ByVal Value As String): m_Language = Value: End Property
Public Property Let ParentalRating(ByVal Value As String): m_ParentalRating = Value: End Property
Public Property Let EncodingBitrate(ByVal Value As String): m_EncodingBitrate = Value: End Property
Public Property Let ChannelCount(ByVal Value As String): m_ChannelCount = Value: End Property
Public Property Let PeakValue(ByVal Value As String): m_PeakValue = Value: End Property
Public Property Let SampleRate(ByVal Value As String): m_SampleRate = Value: End Property
Public Property Let SubTitle(ByVal Value As String): m_SubTitle = Value: End Property
Public Property Let ProviderStyle(ByVal Value As String): m_ProviderStyle = Value: End Property
Public Property Let Year(ByVal Value As Variant): m_Year = Value: End Property
Public Property Let DateEncoded(ByVal Value As String): m_DateEncoded = Value: End Property
Public Property Let Writer(ByVal Value As String): m_Writer = Value: End Property
Public Property Let AverageLevel(ByVal Value As String): m_AverageLevel = Value: End Property
Public Property Let Duration(ByVal Value As String): m_Duration = Value: End Property
Public Property Let Publisher(ByVal Value As String): m_Publisher = Value: End Property
Public Property Let EncodedBy(ByVal Value As String): m_EncodedBy = Value: End Property
Public Property Let AlbumTitle(ByVal Value As String): m_AlbumTitle = Value: End Property
Public Property Let Artist(ByVal Value As Variant): m_Artist = Value: End Property
Public Property Let Composer(ByVal Value As Variant): m_Composer = Value: End Property
Public Property Let Conductor(ByVal Value As String): m_Conductor = Value: End Property
Public Property Let Lyrics(ByVal Value As String): m_Lyrics = Value: End Property
Public Property Let Mood(ByVal Value As String): m_Mood = Value: End Property
Public Property Let TrackNumber(ByVal Value As String): m_TrackNumber = Value: End Property
Public Property Let Genre(ByVal Value As Variant): m_Genre = Value: End Property
Public Property Let AlbumArtist(ByVal Value As String): m_AlbumArtist = Value: End Property


