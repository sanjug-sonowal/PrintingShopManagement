VERSION 5.00
Begin VB.Form AboutForm1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11190
   ScaleWidth      =   21585
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Index           =   3
      Left            =   5880
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Left            =   16200
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   10680
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   4560
      Top             =   3960
   End
   Begin ShopManagementSystem.LabelPlus lblData2 
      Height          =   975
      Index           =   3
      Left            =   7320
      TabIndex        =   13
      Top             =   2760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1720
      BackColor       =   4210752
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "AboutForm1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData2 
      Height          =   3135
      Index           =   2
      Left            =   15240
      TabIndex        =   12
      Top             =   5880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
      BackColor       =   4210752
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "AboutForm1.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData2 
      Height          =   2895
      Index           =   1
      Left            =   9600
      TabIndex        =   11
      Top             =   5880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5106
      BackColor       =   4210752
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "AboutForm1.frx":0040
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData2 
      Height          =   3135
      Index           =   0
      Left            =   3600
      TabIndex        =   10
      Top             =   5640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
      BackColor       =   4210752
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "AboutForm1.frx":0060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   1095
      Index           =   3
      Left            =   6960
      TabIndex        =   9
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1931
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "AboutForm1.frx":0080
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   2
      Left            =   14760
      TabIndex        =   8
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8916
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      Caption         =   "AboutForm1.frx":00B0
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   1
      Left            =   9120
      TabIndex        =   7
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8916
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      Caption         =   "AboutForm1.frx":00D0
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblData1 
      Height          =   5055
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8916
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      Caption         =   "AboutForm1.frx":00F0
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   2175
      Index           =   3
      Left            =   6360
      TabIndex        =   5
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3836
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   6375
      Index           =   2
      Left            =   14040
      TabIndex        =   4
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11245
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   6375
      Index           =   1
      Left            =   8400
      TabIndex        =   3
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11245
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   6375
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11245
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   10815
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   21735
      _ExtentX        =   38338
      _ExtentY        =   19076
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   11895
      Left            =   -360
      TabIndex        =   0
      Top             =   -360
      Width           =   22335
      _ExtentX        =   39396
      _ExtentY        =   20981
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Image Image1 
      Height          =   11175
      Left            =   0
      Picture         =   "AboutForm1.frx":0110
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21615
   End
End
Attribute VB_Name = "AboutForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Dim mFont2 As StdFont
Option Explicit

Private Sub Form_Load()

AboutForm1.Move 0, 0, Me.Width, Me.Height
    
    
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
    Set mFont2 = New StdFont
    mFont2.Name = "Segoe UI"
    mFont2.Size = 10
    mFont2.Bold = True
    
End Sub
Private Sub Form_Resize()
    AboutForm1.Width = Me.ScaleWidth
    AboutForm1.Height = Me.ScaleHeight
    
    
    ResizeControls Me, initialcontrollist
End Sub

Private Sub lblData1_MouseEnter(Index As Integer)
    Timer1(Index).Tag = 1
    Timer1(Index).Interval = 10
End Sub

Private Sub lblData1_MouseLeave(Index As Integer)
    Timer1(Index).Tag = -1
    Timer1(Index).Interval = 10
End Sub

Private Sub lblData2_PostPaint(Index As Integer, ByVal hdc As Long)
    Dim mTop As Long, TextHeight As Long
    Dim sTitle As String
    Dim bProtected As Boolean
    Dim cProtected As Boolean
    Dim dProtected As Boolean
    Dim sDescription As String
    Dim lWidth As Long
    Dim lMargin As Long

    With lblData2(Index)
        
        mTop = 100 - .BackColorOpacity / 1.5

        Select Case Index
        Case 0
            sTitle = "SANJUG SONOWAL"
            bProtected = True
            sDescription = "My Name is Sanjug Jitendra Sonowal and I Am a Professional | Graphic Designer | Software Developer | Freelancer | Pursuing B.C.A From SSR College of Arts, Commerce and Science"
        Case 1
            sTitle = "PRINTING SHOP MANAGEMENT"
            cProtected = True
            sDescription = "This Software is Based on Printing Shop Management , This Software has Many Modern Features and Modern Design. In This Software we Have Integereted Modern Design Calculator, It Has Advance Login System with many Capabilities"
        Case 2
            sTitle = "KALPESH PAWAR"
            bProtected = True
            sDescription = "My Name is Kalpesh Adhaar Pawar , Pursuing B.C.A From SSR College of Arts, Commerce and Science"
        Case 3
            sTitle = "ABOUT SOFTWARE AND DEVELOPERS"
            dProtected = True
            sDescription = "You Can Hover Over the Given Below Cards to Know More Details About our Software Product and Our Developers"
        
        End Select
                
        lMargin = 10 '* .GetWindowsDPI0
        lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)
                                                                  '100= aproximate height
        TextHeight = .DrawText(hdc, sTitle, lMargin, mTop, lWidth, 100, .Font, vbWhite, 100, ccEnter, cTop, True)
        
        If bProtected Then
            TextHeight = TextHeight + .DrawText(hdc, "9724224417", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H88CC44, 100, ccEnter, cTop, True)
            
         ElseIf cProtected Then
            TextHeight = TextHeight + .DrawText(hdc, "9724224417", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H88CC44, 100, ccEnter, cTop, True)
            TextHeight = TextHeight + .DrawText(hdc, "7874352707", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H88CC44, 100, ccEnter, cTop, True)
            
         ElseIf dProtected Then
            TextHeight = TextHeight + .DrawText(hdc, "SANJUG AND KALPESH", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H88CC44, 100, ccEnter, cTop, True)
             
        Else
            TextHeight = TextHeight + .DrawText(hdc, "7874352707", lMargin, mTop + TextHeight, lWidth, 100, mFont2, &HCCC1BB, 100, ccEnter, cTop, True)
        End If
   
        If .BackColorOpacity > 20 Then
            .DrawText hdc, sDescription, lMargin, mTop + TextHeight, lWidth, 200, mFont2, &HCCC1BB, 100, ccEnter, cTop, True
        End If
    End With
End Sub

Private Sub Timer1_Timer(Index As Integer)

    With lblData2(Index)
        .BackColorOpacity = .BackColorOpacity + (10 * Timer1(Index).Tag)
        If .BackColorOpacity = 100 Or .BackColorOpacity = 0 Then Timer1(Index).Interval = 0
    End With

End Sub
