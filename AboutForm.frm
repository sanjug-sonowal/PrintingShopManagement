VERSION 5.00
Begin VB.Form AboutForm 
   Caption         =   "About Form"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21150
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   21150
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Index           =   3
      Left            =   6120
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Left            =   15960
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   10320
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   4080
      Top             =   3480
   End
   Begin ShopManagementSystem.LabelPlus lblData 
      Height          =   975
      Index           =   3
      Left            =   6720
      TabIndex        =   13
      Top             =   1200
      Width           =   6615
      _extentx        =   11668
      _extenty        =   1720
      backcolor       =   4210752
      backcoloropacity=   0
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":0000
      captionbordercolor=   0
      font            =   "AboutForm.frx":0020
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":004C
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData 
      Height          =   5055
      Index           =   2
      Left            =   14280
      TabIndex        =   12
      Top             =   4560
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      backcoloropacity=   0
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":0078
      captionbordercolor=   0
      font            =   "AboutForm.frx":0098
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":00C4
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData 
      Height          =   5055
      Index           =   1
      Left            =   8520
      TabIndex        =   11
      Top             =   4440
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      backcoloropacity=   0
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":00F0
      captionbordercolor=   0
      font            =   "AboutForm.frx":0110
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":013C
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData 
      Height          =   5055
      Index           =   0
      Left            =   2520
      TabIndex        =   10
      Top             =   4440
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      backcoloropacity=   0
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":0168
      captionbordercolor=   0
      font            =   "AboutForm.frx":0188
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":01B4
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   975
      Index           =   3
      Left            =   6720
      TabIndex        =   9
      Top             =   1200
      Width           =   6735
      _extentx        =   11880
      _extenty        =   1720
      backcolor       =   4210752
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":01E0
      captionbordercolor=   0
      font            =   "AboutForm.frx":0200
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":022C
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   2
      Left            =   14280
      TabIndex        =   8
      Top             =   4800
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":0258
      captionbordercolor=   0
      font            =   "AboutForm.frx":0278
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":02A4
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   1
      Left            =   8520
      TabIndex        =   7
      Top             =   4680
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":02D0
      captionbordercolor=   0
      font            =   "AboutForm.frx":02F0
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":031C
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   4680
      Width           =   4095
      _extentx        =   7223
      _extenty        =   8916
      backcolor       =   4210752
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      caption         =   "AboutForm.frx":0348
      captionbordercolor=   0
      font            =   "AboutForm.frx":0368
      picturealignmenth=   1
      picturepaddingy =   30
      shadowsize      =   7
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "AboutForm.frx":0394
      iconforecolor   =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism6 
      Height          =   1815
      Left            =   6240
      TabIndex        =   5
      Top             =   720
      Width           =   7695
      _extentx        =   13573
      _extenty        =   3201
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism5 
      Height          =   6495
      Left            =   13680
      TabIndex        =   4
      Top             =   3960
      Width           =   5295
      _extentx        =   9340
      _extenty        =   11456
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism4 
      Height          =   6495
      Left            =   7920
      TabIndex        =   3
      Top             =   3960
      Width           =   5295
      _extentx        =   9340
      _extenty        =   11456
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   6495
      Left            =   1920
      TabIndex        =   2
      Top             =   3960
      Width           =   5295
      _extentx        =   9340
      _extenty        =   11456
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   10695
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   20655
      _extentx        =   36433
      _extenty        =   18865
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   12015
      Left            =   -360
      TabIndex        =   0
      Top             =   -480
      Width           =   21855
      _extentx        =   38550
      _extenty        =   21193
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin VB.Image Image1 
      Height          =   11055
      Left            =   120
      Picture         =   "AboutForm.frx":03C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20895
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------'
'This Code is for making responsive form

Private initialcontrollist() As ControlInitial
'---------------------------------------------------------------------------------------------------------------------------------'
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Dim mFont2 As StdFont

Private Sub Form_Load()
    Image1.Move 0, 0, Me.Width, Me.Height
   AboutForm.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub
'--------------------------------------------------------------------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------------------------------------------'
Private Sub Form_Resize()
Set mFont2 = New StdFont
    mFont2.Name = "Segoe UI"
    mFont2.Size = 10
    Image1.Width = Me.ScaleWidth
    Image1.Height = Me.ScaleHeight
    
    AboutForm.Width = Me.ScaleWidth
    AboutForm.Height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub

Private Sub lblData_PostPaint(Index As Integer, ByVal hdc As Long)
    Dim mTop As Long, TextHeight As Long
    Dim sTitle As String
    Dim bProtected As Boolean
    Dim sDescription As String
    Dim lWidth As Long
    Dim lMargin As Long

    With lblData(Index)
        
        mTop = 100 - .BackColorOpacity / 1.5

        Select Case Index
        Case 0
            sTitle = "SANJUG SONOWAL"
            bProtected = True
            sDescription = "Name is Sanjug Jitendra Sonowal Pursuing B.C.A From SSR COLLEGE OF ARTS COMMERCE AND SCIENCE"
        Case 1
            sTitle = "PRINTING SHOP MANAGEMENT SYSTEM"
            bProtected = True
            sDescription = "Bloquee amenazas que provengan de la web o correo electónico"
        Case 2
            sTitle = "KALPESH PAWAR"
            bProtected = False
            sDescription = "bca student"
        Case 3
            sTitle = "ABOUT SOFTWARE AND DEVELOPERS"
            bProtected = True
            sDescription = "Protéjase de frisgones"
        
        End Select
                
        lMargin = 10 '* .GetWindowsDPI0
        lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)
                                                                  '100= aproximate height
        TextHeight = .DrawText(hdc, sTitle, lMargin, mTop, lWidth, 100, .Font, vbWhite, 100, ccEnter, cTop, True)
        
        If bProtected Then
            TextHeight = TextHeight + .DrawText(hdc, "Protegido", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H88CC44, 100, ccEnter, cTop, True)
        Else
            TextHeight = TextHeight + .DrawText(hdc, "Sin Protección", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &HCCC1BB, 100, ccEnter, cTop, True)
        End If
   
        If .BackColorOpacity > 20 Then
            .DrawText hdc, sDescription, lMargin, mTop + TextHeight, lWidth, 200, mFont2, &HCCC1BB, 100, ccEnter, cTop, True
        End If
    End With
End Sub


Private Sub lblData1_MouseEnter(Index As Integer)
    Timer1(Index).Tag = 1
    Timer1(Index).Interval = 10
End Sub

Private Sub lblData1_MouseLeave(Index As Integer)
    Timer1(Index).Tag = -1
    Timer1(Index).Interval = 10
End Sub

Private Sub Timer1_Timer(Index As Integer)
 With lblData(Index)
        .BackColorOpacity = .BackColorOpacity + (10 * Timer1(Index).Tag)
        If .BackColorOpacity = 100 Or .BackColorOpacity = 0 Then Timer1(Index).Interval = 0
    End With
End Sub
