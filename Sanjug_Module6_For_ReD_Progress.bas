Attribute VB_Name = "Sanjug_Module6_for_red_progress_state"
Option Explicit

Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Public Const WM_COMMAND = &H111
Private Const WM_DESTROY = &H2

Public Function SubclassForm(ByRef Form As VB.Form) As Boolean
    SubclassForm = SetWindowSubclass(Form.hwnd, AddressOf FrmSubclassProc, ObjPtr(Form)):      Debug.Assert SubclassForm
End Function

Public Function UnSubclassForm(ByRef Form As VB.Form) As Boolean
    UnSubclassForm = RemoveWindowSubclass(Form.hwnd, AddressOf FrmSubclassProc, ObjPtr(Form)): Debug.Assert UnSubclassForm
End Function

Public Function Subclass(hwnd As Long, lpfn As Long, Optional uID As Long = 0&, Optional dwRefData As Long = 0&) As Boolean
If uID = 0 Then uID = hwnd
    Subclass = SetWindowSubclass(hwnd, lpfn, uID, dwRefData):      Debug.Assert Subclass
End Function

Public Function UnSubclass(hwnd As Long, ByVal lpfn As Long, pID As Long) As Boolean
    UnSubclass = RemoveWindowSubclass(hwnd, lpfn, pID)
End Function
Private Function FrmSubclassProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As ChangePassword, ByVal dwRefData As Long) As Long
    Select Case uMsg
        Case WM_COMMAND
            dwRefData = LoWord_B(wParam)  'Avoid declaring additional variables inside window procedures!
            Select Case dwRefData       'They'll still be allocated even for messages not handled here!
                Case 100, 101, 102
                    
                    Exit Function
            End Select

        Case WM_DESTROY
            UnSubclassForm uIdSubclass
    End Select

    FrmSubclassProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
End Function
Public Function F1WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
On Error GoTo e0
  Select Case uMsg

    Case WM_COMMAND
        Dim lw As Integer
        lw = LoWord_B(wParam)
        
    Case WM_DESTROY
      Call UnSubclass(hwnd, PtrF1WndProc, uIdSubclass)

  End Select
  
  F1WndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
Exit Function
e0:
Debug.Print "F1WndProc.Error->" & Err.Description
End Function
Private Function PtrF1WndProc() As Long
PtrF1WndProc = FARPROC(AddressOf F1WndProc)
End Function
Private Function FARPROC(ByVal lpfn As Long) As Long
FARPROC = lpfn
End Function
Public Function LoWord_B(ByVal LongVal As Long) As Integer
    LoWord_B = LongVal And &HFFFF&
End Function


