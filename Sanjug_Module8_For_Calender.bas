Attribute VB_Name = "Sanjug_Module8_For_Calender"
Option Explicit
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const MA_NOACTIVATE As Long = 3
Private Const MA_NOACTIVATEANDEAT As Long = 4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const SWP_NOACTIVATE As Long = &H10
'Private Const SWP_NOMOVE As Long = &H2
'Private Const SWP_FRAMECHANGED As Long = &H20
'Private Const SWP_NOZORDER As Long = &H4
'Private Const SWP_NOOWNERZORDER As Long = &H200
'Private Const SWP_NOSIZE As Long = &H1
'Private Const SWP_SHOWWINDOW As Long = &H40

'Private Const WM_MOUSEACTIVATE As Long = &H21
'Private Const MA_NOACTIVATE As Long = 3
'Private Const MA_NOACTIVATEANDEAT As Long = 4
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Public Type KBDLLHOOKSTRUCT
    VkCode                  As Long
    ScanCode                As Long
    flags                   As Long
    Time                    As Long
    DwExtraInfo             As Long
End Type
 
Dim KBHook                  As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_DESTROY As Long = &H2
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KEYDOWN As Long = &H100
Private Const WH_KEYBOARD_LL   As Long = 13
Dim hHook As Long
Dim hPrevParentProc As Long
Dim m_Hwnd As Long


Public Sub StartHook(hwnd As Long, hWndParent As Long)
    
    m_Hwnd = hwnd
    
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_CHILD
    
    hPrevParentProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf ParentWinProc)
    KBHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KBProc, App.hInstance, 0)
    Do While IsWindowVisible(hwnd)
        DoEvents
    Loop
    SetWindowLong hWndParent, GWL_WNDPROC, hPrevParentProc
    Call UnhookWindowsHookEx(KBHook)
    

End Sub

Public Sub StopHook(hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, hHook
End Sub

Private Function ParentWinProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ParentWinProc = CallWindowProc(hPrevParentProc, hwnd, uMsg, wParam, lParam)

    If uMsg = WM_SYSCOMMAND And wParam = SC_CLOSE Then
        ShowWindow m_Hwnd, 0
        Exit Function
    End If


    If uMsg = WM_RBUTTONDOWN Or uMsg = WM_LBUTTONDOWN Then
       ShowWindow m_Hwnd, 0
       Exit Function
    End If
    
    
    If uMsg <> 123 And _
        uMsg <> 124 And _
        uMsg <> 125 And _
        uMsg <> 60 And _
        uMsg <> 174 And _
        uMsg <> 132 And _
        uMsg <> 512 And _
        uMsg <> 127 And _
        uMsg <> 70 And _
        uMsg <> 32 And _
        uMsg <> 160 And _
        uMsg <> 674 And _
        uMsg <> 134 And _
        uMsg <> 514 And _
        uMsg <> 533 And _
        uMsg <> 517 And _
        uMsg <> 13 And _
        uMsg <> 14 And _
        uMsg <> 15 And _
        uMsg <> 20 And _
        uMsg <> 307 And _
        uMsg <> 261 And _
        uMsg <> 78 Then
        'Debug.Print uMsg
            ShowWindow m_Hwnd, 0
    End If




End Function


 
Public Function KBProc(ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long
    Dim KeyBoardHook        As KBDLLHOOKSTRUCT
 
    If nCode = 0 Then
        CopyMemory KeyBoardHook, lParam, Len(KeyBoardHook)
        With KeyBoardHook
            If .flags = 0 Or .flags = 1 Then
                'Debug.Print .VkCode

                'Select Case .VkCode
                '    Case 37 To 40
                        KBProc = 1
                        SendMessage m_Hwnd, WM_KEYDOWN, .VkCode, 0&
                        Exit Function
                '    Case 27
                '        KBProc = 1
                '        ShowWindow m_Hwnd, 0
                 '       Exit Function
                'End Select
                'CopyMemory lParam, KeyBoardHook, Len(KeyBoardHook)

            End If
        End With
    Else
        KBProc = CallNextHookEx(KBHook, nCode, wParam, lParam)
    End If
End Function


