Attribute VB_Name = "Sanjug_Module1_For_Icon"
Option Explicit

'This code was mostly written by Leandro Ascierto, from his clsMenuImage.
'I've simply modified the resource->hicon function to stand alone
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type IconHeader
    ihReserved      As Integer
    ihType          As Integer
    ihCount         As Integer
End Type

Private Type IconEntry
    ieWidth         As Byte
    ieHeight        As Byte
    ieColorCount    As Byte
    ieReserved      As Byte
    iePlanes        As Integer
    ieBitCount      As Integer
    ieBytesInRes    As Long
    ieImageOffset   As Long
End Type

Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, _
                                                                    ByVal dwResSize As Long, _
                                                                    ByVal fIcon As Long, _
                                                                    ByVal dwVer As Long, _
                                                                    ByVal cxDesired As Long, _
                                                                    ByVal cyDesired As Long, _
                                                                    ByVal Flags As Long) As Long
Private Declare Function CreateIconFromResource Lib "user32.dll" (ByVal presbits As Long, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Private Declare Function LookupIconIdFromDirectoryEx Lib "user32.dll" (ByVal presbits As Long, ByVal fIcon As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long



Public Function ResIconToHICON(id As String, Optional cx As Long = 24, Optional cy As Long = 24) As Long
'returns an hIcon from an icon in the resource file
'Icons must be added as a custom resource

    Dim tIconHeader     As IconHeader
    Dim tIconEntry()    As IconEntry
    Dim MaxBitCount     As Long
    Dim MaxSize         As Long
    Dim Aproximate      As Long
    Dim IconID          As Long
    Dim hIcon           As Long
    Dim i               As Long
    Dim bytIcoData() As Byte
    
On Error GoTo e0

    bytIcoData = LoadResData(id, "CUSTOM")
    Call CopyMemory(tIconHeader, bytIcoData(0), Len(tIconHeader))

    If tIconHeader.ihCount >= 1 Then
    
        ReDim tIconEntry(tIconHeader.ihCount - 1)
        
        Call CopyMemory(tIconEntry(0), bytIcoData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
        
        IconID = -1
           
        For i = 0 To tIconHeader.ihCount - 1
            If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
        Next

       
        For i = 0 To tIconHeader.ihCount - 1
            If MaxBitCount = tIconEntry(i).ieBitCount Then
                MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                If MaxSize > Aproximate And MaxSize <= (cx + cy) Then
                    Aproximate = MaxSize
                    IconID = i
                End If
            End If
        Next
                   
        If IconID = -1 Then Exit Function
       
        With tIconEntry(IconID)
            hIcon = CreateIconFromResourceEx(bytIcoData(.ieImageOffset), .ieBytesInRes, 1, &H30000, cx, cy, &H0)
            If hIcon <> 0 Then
                ResIconToHICON = hIcon
            End If
        End With
       
    End If
'Debug.Print "Res hIcon=" & hIcon

On Error GoTo 0
Exit Function

e0:
Debug.Print "modIcon.ResIconTohIcon.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Public Function IconToHICON(IcoData() As Byte, DesiredX As Long, DesiredY As Long) As Long
    Dim lPtrSrc As Long, lPtrDst As Long, lID As Long
    Dim icDir() As Byte, LB As Long
    Dim tIconHeader As IconHeader
    Dim tIconEntry As IconEntry
    Dim ICRESVER As Long
    ICRESVER = &H30000
    LB = LBound(IcoData) ' just in case a non-zero LBound array passed
    ' convert 16 byte IconDir to 14 byte IconDir
    CopyMemory tIconHeader, IcoData(LB), Len(tIconHeader)
    ReDim icDir(0 To tIconHeader.ihCount * Len(tIconEntry) + Len(tIconHeader) - 1&)
    CopyMemory icDir(0), tIconHeader, Len(tIconHeader)
    lPtrDst = Len(tIconHeader)
    lPtrSrc = LB + lPtrDst
    For lID = 1& To tIconHeader.ihCount
        CopyMemory tIconEntry, IcoData(lPtrSrc), 12& ' size of standard tIconEntry less last 4 bytes
        tIconEntry.ieImageOffset = lID
        CopyMemory icDir(lPtrDst), tIconEntry, 14&     ' size of DLL tIconEntry
        lPtrDst = lPtrDst + 14&: lPtrSrc = lPtrSrc + Len(tIconEntry)
    Next
    lID = LookupIconIdFromDirectoryEx(VarPtr(icDir(0)), True, DesiredX, DesiredY, 0&)
    Erase icDir()
    If lID > 0& Then
        CopyMemory tIconEntry, IcoData(LB + (lID - 1&) * Len(tIconEntry) + Len(tIconHeader)), Len(tIconEntry)
        
        IconToHICON = CreateIconFromResource(VarPtr(IcoData(LB + tIconEntry.ieImageOffset)), tIconEntry.ieBytesInRes, True, ICRESVER)
    End If
End Function
Public Function LoadIcoFile(sFile As String) As Byte()
    Dim f As Long
    'Dim b() As Byte
     
    f = FreeFile()
    Open sFile For Binary As f
    ReDim LoadIcoFile(LOF(f))
    Get f, , LoadIcoFile
    Close f
End Function

