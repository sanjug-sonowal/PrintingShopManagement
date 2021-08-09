Attribute VB_Name = "MdlUnicodeFunctions"
Option Explicit

Private Const MAX_PATH = 260

Public Enum VbFileAttributeExtended
    vbAll = -1&
    vbDirectory = 16& ' mean - include folders also
    vbFile = vbAll And Not vbDirectory
    vbSystem = 4&
    vbHidden = 2&
    vbReadOnly = 1
    vbNormal = 0&
    vbReparse = 1024& 'symlinks / junctions (not include hardlink to file; they reflect attributes of the target)
End Enum


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH) As Integer
    lpszAlternate(14) As Integer
End Type

Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1
 

Const MAX_PATH_W    As Long = 32767&
 
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
 
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As OSVERSIONINFOEX) As Long
 
Private Const VER_NT_WORKSTATION        As Long = 1&
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function CommandLineToArgvW Lib "Shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetCurrentDirectory Lib "kernel32.dll" Alias "SetCurrentDirectoryW" (ByVal lpPathName As Long) As Long


Public Function ParsedCommandLine(Out() As String) As Boolean
    Dim ptr     As Long
    Dim Count   As Long
    Dim Index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    ptr = CommandLineToArgvW(GetCommandLine(), Count)
    If Count < 1 Then Exit Function
    ReDim Out(Count - 1)
    For Index = 0 To Count - 1
        GetMem4 ByVal ptr + Index * 4, strAdr
        strLen = lstrlen(strAdr)
        Out(Index) = Space(strLen)
        lstrcpyn StrPtr(Out(Index)), strAdr, strLen + 1
    Next
    GlobalFree ptr
    ParsedCommandLine = True
End Function

Public Sub ChDirW(ByVal sPath As String)
    Call SetCurrentDirectory(StrPtr(sPath))
End Sub

Public Function DirW( _
    Optional ByVal PathMaskOrFolderWithSlash As String, _
    Optional AllowedAttributes As VbFileAttributeExtended = vbNormal, _
    Optional FoldersOnly As Boolean) As String

    On Error GoTo ErrorHandler

    'WARNING note:
    'Original VB Dir$ contains bug: ReadOnly attribute incorrectly handled, so it always is in results
    'This sub properly handles 'RO' and also contains one extra flag: FILE_ATTRIBUTE_REPARSE_POINT (vbReparse)
    'Doesn't return "." and ".." folders.
    'Unicode aware

    Const MeaningfulBits As Long = &H417&   'D + H + R + S + Reparse
                                            '(to revert to default VB Dir behaviour, replace it by &H16 value)

    Dim fd      As WIN32_FIND_DATA
    Dim lpStr   As Long
    Dim lRet    As Long
    Dim Mask    As Long

    Static hFind        As Long
    Static lflags       As VbFileAttributeExtended
    Static bFoldersOnly As Boolean

    If hFind <> 0& And Len(PathMaskOrFolderWithSlash) = 0& Then
        If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0&: Exit Function
    Else
        If hFind Then FindClose hFind: hFind = 0&
        PathMaskOrFolderWithSlash = Trim(PathMaskOrFolderWithSlash)
        lflags = AllowedAttributes 'cache
        bFoldersOnly = FoldersOnly 'cache

        Select Case Right$(PathMaskOrFolderWithSlash, 1&)
        Case "\", ":", "/"
            PathMaskOrFolderWithSlash = PathMaskOrFolderWithSlash & "*.*"
        End Select

        hFind = FindFirstFile(StrPtr(PathMaskOrFolderWithSlash), fd)

        If hFind = INVALID_HANDLE_VALUE Then
            If (err.LastDllError) > 12& Then hFind = 0&: err.Raise 52&
            Exit Function
        End If
    End If

    Do
        If fd.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
            Mask = 0& 'found
        Else
            Mask = fd.dwFileAttributes And (Not lflags) And MeaningfulBits
        End If
        If bFoldersOnly Then
            If Not CBool(fd.dwFileAttributes And vbDirectory) Then
                Mask = 1 'continue enum
            End If
        End If

        If Mask = 0 Then
            lpStr = VarPtr(fd.lpszFileName(0))
            DirW = String$(lstrlen(lpStr), 0&)
            lstrcpy StrPtr(DirW), lpStr
            If fd.dwFileAttributes And vbDirectory Then
                If DirW <> "." And DirW <> ".." Then Exit Do 'exclude self and relative paths aliases
            Else
                Exit Do
            End If
        End If

        If FindNextFile(hFind, fd) = 0 Then FindClose hFind: hFind = 0: Exit Function
    Loop

    Exit Function
ErrorHandler:
    Debug.Print err; err.Description; "DirW"
End Function

Public Function IsFolder(ByVal Path As String) As Boolean
    Dim L As Long
    Path = Trim(Path)
    L = GetFileAttributes(StrPtr(Path))
    IsFolder = CBool(L And vbDirectory) And (L <> INVALID_HANDLE_VALUE)
End Function

Public Function IsFile(ByVal Path As String) As Boolean
    Dim L As Long
    Path = Trim(Path)
    L = GetFileAttributes(StrPtr(Path))
    IsFile = Not CBool(L And vbDirectory) And (L <> INVALID_HANDLE_VALUE)
End Function

Public Function KillW(ByVal Path As String) As Boolean
   KillW = CBool(DeleteFileW(StrPtr(Path)) <> INVALID_HANDLE_VALUE)
End Function

Function MkDirW(ByVal Path As String, Optional ByVal LastComponentIsFile As Boolean = False) As Long
    Dim FC As String, lr As Long, pos As Long
    If LastComponentIsFile Then Path = Left(Path, InStrRev(Path, "\") - 1)
    If InStr(Path, ":") = 0 Then
        Dim sCurDir$, nChar As Long
        sCurDir = String$(MAX_PATH, 0&)
        nChar = GetCurrentDirectory(MAX_PATH + 1, StrPtr(sCurDir))
        sCurDir = Left$(sCurDir, nChar)
        If Right$(sCurDir, 1) <> "\" Then sCurDir = sCurDir & "\"
        Path = sCurDir & Path
    End If
    Do
        pos = pos + 1
        pos = InStr(pos, Path, "\")
        If pos Then FC = Left(Path, pos - 1) Else FC = Path
        lr = 1
        If Not IsFolder(FC) Then lr = CreateDirectory(StrPtr(FC), ByVal 0&)
    Loop While (pos <> 0) And (lr <> 0)
    MkDirW = lr
End Function

Public Function AppPathW(Optional bGetFullPath As Boolean) As String
    On Error GoTo ErrorHandler
 
    Static ProcPathFull  As String
    Static ProcPathShort As String
    Dim ProcPath As String
    Dim Cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long
    Dim inIDE    As Boolean
    
    'Cache
    If bGetFullPath Then
        If Len(ProcPathFull) <> 0 Then
            AppPathW = ProcPathFull
            Exit Function
        End If
    Else
        If Len(ProcPathShort) <> 0 Then
            AppPathW = ProcPathShort
            Exit Function
        End If
    End If
 
    inIDE = (App.LogMode = 0)
 
    If inIDE Then
        AppPathW = GetDOSFilename(App.Path, bReverse:=True)
        'bGetFullPath does not supported in IDE
        Exit Function
    End If
 
    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0
 
    ProcPath = String$(MAX_PATH, vbNullChar)
    Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If Cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If Cnt = 0 Then                          'clear path
        ProcPath = App.Path
    Else
        ProcPath = Left$(ProcPath, Cnt)
        If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = GetWindowsDir() & Mid$(ProcPath, 12)
        If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        
        If Not bGetFullPath Then
            ' trim to path
            pos = InStrRev(ProcPath, "\")
            If pos <> 0 Then ProcPath = Left$(ProcPath, pos - 1)
        End If
    End If
    
    ProcPath = GetDOSFilename(ProcPath, bReverse:=True)     '8.3 -> to Full
    
    AppPathW = ProcPath
    
    If bGetFullPath Then
        ProcPathFull = ProcPath
    Else
        ProcPathShort = ProcPath
    End If
    Exit Function
ErrorHandler:
    Debug.Print err; "Parser.AppPath"; "ProcPath:"; ProcPath
    'If inIDE Then Stop: Resume Next
End Function
 
Public Function AppExeNameW(Optional WithExtension As Boolean) As String
    On Error GoTo ErrorHandler
 
    Static ProcPathShort As String
    Static ProcPathFull  As String
    Dim ProcPath As String
    Dim Cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long
    Dim inIDE    As Boolean
 
    'Cache
    If WithExtension Then
        If Len(ProcPathFull) <> 0 Then
            AppExeNameW = ProcPathFull
            Exit Function
        End If
    Else
        If Len(ProcPathShort) <> 0 Then
            AppExeNameW = ProcPathShort
            Exit Function
        End If
    End If
 
    inIDE = (App.LogMode = 0)
 
    If inIDE Then
        AppExeNameW = App.EXEName & IIf(WithExtension, ".exe", "")
        Exit Function
    End If
 
    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0
 
    ProcPath = String$(MAX_PATH, vbNullChar)
    Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If Cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        Cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If Cnt = 0 Then                          'clear path
        ProcPath = App.EXEName & IIf(WithExtension, ".exe", "")
    Else
        ProcPath = Left$(ProcPath, Cnt)
        
        pos = InStrRev(ProcPath, "\")
        If pos <> 0 Then ProcPath = Mid$(ProcPath, pos + 1)
        
        If Not WithExtension Then
            ProcPath = GetFileName(ProcPath)
        End If
    End If
    
    AppExeNameW = ProcPath
    
    If WithExtension Then
        ProcPathFull = ProcPath
    Else
        ProcPathShort = ProcPath
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print err; "Parser.AppExeName"; "ProcPath:"; ProcPath
    'If inIDE Then Stop: Resume Next
End Function
 
'if short name is unavailable, it returns source string anyway
Public Function GetDOSFilename$(sFile$, Optional bReverse As Boolean = False)
    'works for folders too btw
    Dim Cnt&, sBuffer$
    If bReverse Then
        sBuffer = Space$(MAX_PATH_W)
        Cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If Cnt Then
            GetDOSFilename = Left$(sBuffer, Cnt)
        Else
            GetDOSFilename = sFile
        End If
    Else
        sBuffer = Space$(MAX_PATH)
        Cnt = GetShortPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If Cnt Then
            GetDOSFilename = Left$(sBuffer, Cnt)
        Else
            GetDOSFilename = sFile
        End If
    End If
End Function
 
Public Function GetWindowsDir() As String
    Static SysRoot As String
    Static IsInit As Boolean
    Dim lr As Long
    Dim osi As OSVERSIONINFOEX
    
    If IsInit Then
        GetWindowsDir = SysRoot
        Exit Function
    End If
    
    IsInit = True
    GetVersionEx osi
    
    If osi.wProductType = VER_NT_WORKSTATION Then
        SysRoot = String$(MAX_PATH, 0&)
        lr = GetWindowsDirectory(StrPtr(SysRoot), MAX_PATH)
        If lr Then
            SysRoot = Left$(SysRoot, lr)
        Else
            SysRoot = Environ$("SystemRoot")
        End If
    Else
        SysRoot = Environ$("SystemRoot") 'avoid path virtualization on Windows Server with Terminal Services
    End If
    GetWindowsDir = SysRoot
End Function
 
' Get only file name (without extension)
Public Function GetFileName(Path As String) As String
    Dim posDot      As Long
    Dim posSl       As Long
    
    posSl = InStrRev(Path, "\")
    If posSl <> 0 Then
        posDot = InStrRev(Path, ".")
        If posDot < posSl Then posDot = 0
    Else
        posDot = InStrRev(Path, ".")
    End If
    If posDot = 0 Then posDot = Len(Path) + 1
    
    GetFileName = Mid$(Path, posSl + 1, posDot - posSl - 1)
End Function

Public Function GetFileExt(Path As String) As String
    Dim pos      As Long

    pos = InStrRev(Path, ".")
    If pos <> 0 Then
        GetFileExt = LCase$(Trim$(Mid$(Path, pos)))
    End If
End Function

