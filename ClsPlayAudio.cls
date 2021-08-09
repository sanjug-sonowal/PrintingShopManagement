VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPlayAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FilePath As String
Private m_ChannelLength As Double
Public chan As Long
Private m_Pause As Boolean
Private WithEvents cTimer As ClsTimer
Attribute cTimer.VB_VarHelpID = -1
Private WithEvents cTimerRewindAdvance As ClsTimer
Attribute cTimerRewindAdvance.VB_VarHelpID = -1
Public Event Progress(percent As Long, ChannelPosition As Double)
Public Event EndAudio()
Private m_ChannelPosition As Double
Private m_Rewind As Boolean
Private m_Advance As Boolean
Private m_Repeat As Boolean
Private m_OldVolumen As Single
Private m_Mute As Boolean


Public Function Init(ByVal hwnd As Long) As Boolean
 

    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then Exit Function
    If BASS_Init(-1, 44100, 0, hwnd, 0) = BASSFALSE Then Exit Function
    'Debug.Print BASS_Init(-1, 44100, 0, hwnd, 0)

    Set cTimer = New ClsTimer
    Set cTimerRewindAdvance = New ClsTimer
    Init = True
End Function

Public Property Get Mute() As Boolean
    Mute = m_Mute
End Property

Public Property Let Mute(Value As Boolean)
    If Value Then
        If m_Mute = False Then
            Call BASS_ChannelGetAttribute(chan, BASS_ATTRIB_VOL, m_OldVolumen)
            Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_VOL, 0)
            m_Mute = True
        End If
    Else
        If m_Mute Then
            Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_VOL, m_OldVolumen)
            m_Mute = False
        End If
    End If
End Property

 
Public Property Get Volumen() As Long
    Call BASS_ChannelGetAttribute(chan, BASS_ATTRIB_VOL, m_OldVolumen)
    Volumen = m_OldVolumen * 100
End Property

Public Property Let Volumen(Value As Long)
    m_Mute = False
    m_OldVolumen = Value / 100
    Call BASS_ChannelSetAttribute(chan, BASS_ATTRIB_VOL, m_OldVolumen)
End Property

Public Function LoadFile(sPath As String) As Boolean
    On Error GoTo ErrHandler
    If chan <> 0 Then Call BASS_StreamFree(chan)
    
    m_FilePath = vbNullString
    
    If Len(Dir(sPath, vbArchive)) = 0 Then Exit Function
    
    chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(sPath), 0, 0, BASS_STREAM_DECODE)

    If (chan = 0) Then GoTo ErrHandler
    
    m_FilePath = sPath
    m_Pause = True
    m_ChannelLength = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetLength(chan, BASS_POS_BYTE))

    LoadFile = m_ChannelLength > 0
    Exit Function
ErrHandler:
    m_ChannelLength = 0
    Debug.Print "ClsPlayAudio Error LoadFile"
End Function

Public Sub ReleaseFile()
    On Error Resume Next
    StopAudio
    If chan Then Call BASS_StreamFree(chan): chan = 0
    m_ChannelLength = 0
    m_FilePath = vbNullString
    m_Mute = False
End Sub

Public Function Play() As Boolean
    If m_ChannelLength = 0 Then Exit Function
    m_Pause = False
    Play = BASS_ChannelPlay(chan, BASSFALSE)
    If Play Then
        cTimer.CreateTimer 100
    End If
End Function

Public Property Get AudioLength() As Double
    AudioLength = m_ChannelLength
End Property

Public Function StopAudio() As Boolean
    Me.Advance = False
    cTimer.DestroyTimer
    Call BASS_ChannelStop(chan)
    m_Pause = True
    BASS_ChannelSetPosition chan, BASS_ChannelSeconds2Bytes(chan, 0), BASS_POS_BYTE
    RaiseEvent Progress(0, 0)
    
End Function

Public Property Let Repeat(Value As Boolean)
    m_Repeat = Value
End Property

Public Property Get Repeat() As Boolean
    Repeat = m_Repeat
End Property


Public Property Let Pause(Value As Boolean)
    If m_ChannelLength = 0 Then Exit Property
    If Value Then
        Call BASS_ChannelPause(chan)
        cTimer.DestroyTimer
    Else
        Call BASS_ChannelPlay(chan, BASSFALSE)
        cTimer.CreateTimer 100
    End If
    m_Pause = Value
End Property

Public Property Get Pause() As Boolean
    Pause = m_Pause
End Property

Public Property Get isAudioPlay() As Boolean
    isAudioPlay = BASS_ChannelIsActive(chan)
End Property

Public Property Get handle() As Long
    handle = chan
End Property

Public Property Get ChannelPosition() As Double
    ChannelPosition = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))
End Property

Public Property Let ChannelPosition(Value As Double)
    BASS_ChannelSetPosition chan, BASS_ChannelSeconds2Bytes(chan, Value), BASS_POS_BYTE
End Property

Public Property Let ChannelPositionPercent(Value As Long)
    Dim lPercent As Long
    lPercent = m_ChannelLength * Value / 100
    BASS_ChannelSetPosition chan, BASS_ChannelSeconds2Bytes(chan, lPercent), BASS_POS_BYTE
End Property

Public Property Get ChannelPositionPercent() As Long
    m_ChannelPosition = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))
    ChannelPositionPercent = m_ChannelPosition * 100 / m_ChannelLength
End Property

Public Property Let Advance(Value As Boolean)
    m_Rewind = False
    m_Advance = Value
    If Value Then
        cTimerRewindAdvance.CreateTimer 500
    Else
        cTimerRewindAdvance.DestroyTimer
    End If
End Property

Public Property Get Advance() As Boolean
    Advance = m_Advance
End Property

Public Property Let Rewind(Value As Boolean)
    m_Advance = False
    m_Rewind = Value
    If Value Then
        cTimerRewindAdvance.CreateTimer 500
    Else
        cTimerRewindAdvance.DestroyTimer
    End If
End Property

Public Property Get Rewind() As Boolean
    Rewind = m_Rewind
End Property

Public Property Get GetFilePath() As String
    GetFilePath = m_FilePath
End Property

Private Sub cTimer_Timer(ByVal ThisTime As Long)

    Dim lPercent As Long
    If m_ChannelLength = 0 Then Exit Sub

    m_ChannelPosition = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))

    lPercent = m_ChannelPosition * 100 / m_ChannelLength
    RaiseEvent Progress(lPercent, m_ChannelPosition)
    
   
    If BASS_ChannelIsActive(chan) = 0 Then
        Me.StopAudio
        If m_Repeat Then
            Me.Play
        Else
            RaiseEvent EndAudio
        End If
    End If
    
End Sub

Private Sub Class_Terminate()
    Set cTimer = Nothing
    Set cTimerRewindAdvance = Nothing
    If chan Then Call BASS_StreamFree(chan)
    Call BASS_Free
End Sub

Private Sub cTimerRewindAdvance_Timer(ByVal ThisTime As Long)
    If m_Advance Then
        If Me.ChannelPosition >= m_ChannelLength Then
            Me.Advance = False
        Else
            Me.ChannelPosition = Me.ChannelPosition + 3
        End If
        Call cTimer_Timer(0)
    End If
    
    If m_Rewind Then
        If Me.ChannelPosition <= 0 Then
            Me.Rewind = False
        Else
            Me.ChannelPosition = Me.ChannelPosition - 3
        End If
        Call cTimer_Timer(0)
    End If
    
End Sub
