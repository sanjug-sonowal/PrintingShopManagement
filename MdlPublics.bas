Attribute VB_Name = "MdlPublics"
Option Explicit
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal lngHandler As Long, ByVal lngIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal lngHandler As Long, ByVal lngIndex As Long, ByVal lngNewClassLong As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_DESTROY   As Long = &H2
Private Const WM_DROPFILES As Long = &H233


Public Const CS_DROPSHADOW As Long = &H20000
Public Const GCL_STYLE As Long = -26

Private Const WH_KEYBOARD_LL = 13
Private Const HC_ACTION = 0
Private Const HC_NOREMOVE = 3
Private Const SC_CLOSE As Long = &HF060&

Public FDPI As Single

Private hHook As Long
Public bButtonAcept As Boolean
Public ReturText As String
 
Public Sub HookKeyboard()
    hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
End Sub
 
Public Sub RemoveKeyboardHook()
    UnhookWindowsHookEx hHook
End Sub
 
Public Function LowLevelKeyboardProc(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    
    If uCode >= 0 Then
        Select Case uCode
            Case HC_ACTION

                If Form1.Hook_Keyboard(lParam.vkCode, lParam.Flags) Then
                    LowLevelKeyboardProc = 1
                    Exit Function
                End If
 
            Case HC_NOREMOVE
                'The message has not been removed from the message queue
        End Select
    End If
        
    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)
End Function


Public Function SubClass(ByRef Frm As Form) As Boolean
    SubClass = SetWindowSubclass(Frm.hwnd, AddressOf SubclassProc, ObjPtr(Frm), AddressOf SubclassProc) <> 0&
    Debug.Assert SubClass
End Function

Public Function UnSubclass(hwnd As Long, lpfn As Long) As Long
'Only needed if you want to stop the subclassing code and keep the program running.
'Otherwise, the WndProc function should call this on WM_DESTROY
    UnSubclass = RemoveWindowSubclass(hwnd, lpfn, 0)
End Function
Private Function SubclassProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                              ByVal uIdSubclass As Form, ByVal dwRefData As Long) As Long
    Select Case uMsg
        Case WM_DROPFILES:  uIdSubclass.Drag_n_Drop wParam
                            Exit Function
                            
        Case WM_MOUSEWHEEL: uIdSubclass.MouseWhell wParam
        

        Case WM_SYSCOMMAND
            If wParam = SC_CLOSE Then
                Unload uIdSubclass
            End If
            
        Case WM_COMMAND
            dwRefData = LoWord(wParam)
            Select Case dwRefData
                Case 102, 103, 104
                    uIdSubclass.GotButtonClick CInt(dwRefData)
                    Exit Function
            End Select

        
        Case WM_DESTROY:    dwRefData = RemoveWindowSubclass(hwnd, dwRefData, ObjPtr(uIdSubclass))
                            Debug.Assert dwRefData
    End Select

    SubclassProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
End Function

Private Function LoWord(ByVal Numero As Long) As Long
    LoWord = Numero And &HFFFF&
End Function

Private Function HiWord(ByVal Numero As Long) As Long
    HiWord = Numero \ &H10000 And &HFFFF&
End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

