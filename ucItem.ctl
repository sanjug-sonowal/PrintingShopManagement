VERSION 5.00
Begin VB.UserControl ucItem 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit




Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event ButtonClick()
Public Event DblClick()

Dim Margin As Long
Dim m_ButtonState As Boolean


Public Sub SetData(Caption1 As String, Caption2 As String)

    LP1.Caption = Caption1 ': LP1.AutoSize = True
    LP2.Caption = Caption2 ': LP2.AutoSize = True

End Sub

Public Property Get Title() As String
    Title = LP1.Caption
End Property

Public Property Let Title(NewValue As String)
    LP1.Caption = NewValue
End Property

Public Property Get Artist() As String
    Artist = LP2.Caption
End Property

Public Property Let Artist(NewValue As String)
    LP2.Caption = NewValue
End Property

Private Sub N2_Click()
    RaiseEvent ButtonClick
End Sub


Private Sub N2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Let ButtonState(isPlaying As Boolean)
    m_ButtonState = isPlaying
    If isPlaying Then
        LP3.LoadPicture LoadResData("PAUSE", "PNG")
        N2.BackColor = &HC07000
    Else
        LP3.LoadPicture LoadResData("PLAY", "PNG")
        N2.BackColor = &HF2E1D7
    End If
End Property


Public Property Get ButtonState() As Boolean
    ButtonState = m_ButtonState
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub


Private Sub UserControl_Resize()
    If Margin = 0 Then Margin = 5 * LP1.GetWindowsDPI
    N1.Move UserControl.ScaleWidth - N1.Width, UserControl.ScaleHeight / 2 - N1.Height / 2
    N2.Move N1.Left + N1.Width / 2 - N2.Width / 2, UserControl.ScaleHeight / 2 - N2.Height / 2
    LP3.Move N1.Left + N1.Width / 2 - LP3.Width / 2 + 1, UserControl.ScaleHeight / 2 - LP3.Height / 2 + 1
    LP1.Move Margin, UserControl.ScaleHeight / 2 - LP1.Height, UserControl.ScaleWidth - N1.Width - Margin
    LP2.Move Margin, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - N1.Width - Margin
End Sub
