VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBassFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Sub Class_Initialize()
    freq = 44100
End Sub

Public Sub Init()
    ' check the correct BASS_FX was loaded
    If (HiWord(BASS_FX_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of BASS_FX.DLL was loaded (2.4 is required)", vbCritical)
    End If
End Sub

Public Property Let chan(lChan As Long)
    m_Chan = BASS_FX_TempoCreate(lChan, BASS_FX_FREESOURCE)
    Call BASS_ChannelGetInfo(m_Chan, info)
    m_TempoFreq = info.freq
    Call SetDSP_EQ(0, 2.5, 0, 125, 1000, 8000)
End Property

Public Property Get chan() As Long
    chan = m_Chan
End Property

Public Property Let Tempo(NewValue As Single)
    m_Tempo = NewValue
    If m_Chan Then Call BASS_ChannelSetAttribute(m_Chan, BASS_ATTRIB_TEMPO, m_Tempo)
End Property

Public Property Get Tempo() As Single
    Tempo = m_Tempo
End Property

Public Property Let TempoFreq(NewValue As Single)
    'm_OldTempFreq
    m_TempoFreq = NewValue
    If m_Chan Then Call BASS_ChannelSetAttribute(m_Chan, BASS_ATTRIB_TEMPO_FREQ, m_TempoFreq)
End Property

Public Property Get TempoFreq() As Single
    TempoFreq = m_TempoFreq
End Property

Public Property Let TempoPich(NewValue As Single)
    m_TempoPich = NewValue
    If m_Chan Then Call BASS_ChannelSetAttribute(m_Chan, BASS_ATTRIB_TEMPO_PITCH, m_TempoPich)
End Property

Public Property Get TempoPich() As Single
    TempoPich = m_TempoPich
End Property

Public Property Let BandValue(ByVal Index As Integer, ByVal NewValue As Single)
    eq.lBand = Index    ' get values of the selected band
    Call BASS_FXGetParameters(fxEQ, eq)
        eq.fGain = NewValue * -1
    Call BASS_FXSetParameters(fxEQ, eq)
End Property

Public Property Get BandValue(ByVal Index As Integer) As Single
    eq.lBand = Index    ' get values of the selected band
    Call BASS_FXGetParameters(fxEQ, eq)
    BandValue = eq.fGain
End Property

' set dsp peaking eq
Public Sub SetDSP_EQ(ByVal fGain As Single, ByVal fBandwidth As Single, ByVal fQ As Single, ByVal fCenter_Bass As Single, ByVal fCenter_Mid As Single, ByVal fCenter_Treble As Single)
    ' set peaking equalizer effect with no bands
    fxEQ = BASS_ChannelSetFX(m_Chan, BASS_FX_BFX_PEAKEQ, 0)

    eq.fGain = fGain
    eq.fQ = fQ
    eq.fBandwidth = fBandwidth
    eq.lChannel = BASS_BFX_CHANALL

    ' create 1st band for bass
    eq.lBand = 0
    eq.fCenter = fCenter_Bass
    Call BASS_FXSetParameters(fxEQ, eq)

    ' create 2nd band for mid
    eq.lBand = 1
    eq.fCenter = fCenter_Mid
    Call BASS_FXSetParameters(fxEQ, eq)

    ' create 3rd band for treble
    eq.lBand = 2
    eq.fCenter = fCenter_Treble
    Call BASS_FXSetParameters(fxEQ, eq)

End Sub
