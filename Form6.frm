VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerUpdate 
      Left            =   1920
      Top             =   3240
   End
   Begin VB.Frame fraMonth 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   8160
      TabIndex        =   61
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   11
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   73
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   10
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   72
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   8
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   71
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   7
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   70
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   6
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   69
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   4
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   68
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   3
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   67
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   2
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   66
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   0
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   65
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   9
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   64
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   5
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   63
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpMonth 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   1
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   62
         Top             =   90
         Width           =   600
      End
   End
   Begin VB.Frame fraYear 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   4680
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   1
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   60
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   5
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   59
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   9
         Left            =   695
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   58
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   0
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   57
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   2
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   56
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   3
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   55
         Top             =   90
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   4
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   54
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   6
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   53
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   7
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   52
         Top             =   787
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   8
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   51
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   10
         Left            =   1360
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   50
         Top             =   1485
         Width           =   600
      End
      Begin VB.PictureBox LpYear 
         BackColor       =   &H00000000&
         Height          =   600
         Index           =   11
         Left            =   2025
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   49
         Top             =   1485
         Width           =   600
      End
   End
   Begin VB.Frame fraDay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   180
      TabIndex        =   5
      Top             =   795
      Width           =   2655
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   34
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   47
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   33
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   46
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   32
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   45
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   31
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   44
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   30
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   43
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   29
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   42
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   28
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   41
         Top             =   1800
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   27
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   40
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   26
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   39
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   25
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   38
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   24
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   37
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   23
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   36
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   22
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   35
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   21
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   34
         Top             =   1440
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   20
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   33
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   19
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   18
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   31
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   17
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   16
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   29
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   15
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   28
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   14
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   13
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   26
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   12
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   11
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   24
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   10
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   9
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   8
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   6
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox LpNum 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   7
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2220
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1860
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1500
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   780
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   420
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox LpDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox LpLine 
      ForeColor       =   &H00CCCCCC&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.PictureBox LpToday 
      Height          =   360
      Left            =   240
      ScaleHeight     =   300
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   3220
      Width           =   735
   End
   Begin VB.PictureBox LpMonthYear 
      BackColor       =   &H00EEEEEE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.PictureBox LpChangeMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   360
      Width           =   405
   End
   Begin VB.PictureBox LabelPlus1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   0
      Width           =   3060
      Begin ShopManagementSystem.LabelPlus LabelPlus2 
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BackShadow      =   0   'False
         Caption         =   "Form6.frx":0000
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
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ForeColorSelectMonth As Long = &HA78767
Private Const ForeColorOtherMonth As Long = &HCCCCCC
Private Const ThemeColor As Long = &H5F4C49
Private Const ResalteColor As Long = &H993300

Public DateValue As Date

Private m_Index As Long
Private xDate As Date
Private cCLW As cCanvasLayeredWindow


Public Function ShowCalendar(hWndParent As Long) As Boolean
    Dim i As Long
    Dim DPI As Long
    Me.Width = 3135
    m_Index = -1
    DPI = LabelPlus1.GetWindowsDPI
    Set cCLW = New cCanvasLayeredWindow
    Me.ScaleMode = vbPixels
    
    cCLW.CreateCanvas LabelPlus1.Width, LabelPlus1.Height
    
    fraDay.Move 12 * DPI, 53 * DPI
    fraMonth.Move 12 * DPI, 53 * DPI
    fraYear.Move 12 * DPI, 53 * DPI
    
    If Year(DateValue) = 1899 Then DateValue = Now
    xDate = DateValue
    
    For i = 0 To LpDay.Count - 1
      LpDay(i).ForeColor = ThemeColor
      LpDay(i).Caption = StrConv(Left(Format(DateAdd("d", i, FirstDayOfWeek(xDate)), "DDDD"), 2), vbProperCase)
    Next
    
    LpMonthYear.ForeColor = ThemeColor
    LpToday.ForeColor = ThemeColor
    LpToday.BorderColor = ThemeColor
    
    ChangeDate

    StartHook Me.hwnd, hWndParent
    ShowCalendar = m_Index > -1
    Unload Me
End Function

Private Sub Form_Paint()
    Static First As Boolean
    If First = False Then
        Call Update
        First = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyRight
            If m_Index < LpNum.Count - 1 Then
                m_Index = m_Index + 1
            Else
                m_Index = 0
            End If
            
            Changed

        Case vbKeyLeft

            If m_Index > 0 Then
                m_Index = m_Index - 1
            Else
                m_Index = LpNum.Count - 1
            End If
            
            Changed
            
        Case vbKeyDown
            xDate = DateAdd("m", 1, xDate)
            ChangeDate

        Case vbKeyUp
            xDate = DateAdd("m", -1, xDate)
            ChangeDate
            
        Case vbKeyReturn

            If m_Index = -1 Then Exit Sub
            DateValue = LpNum(m_Index).Tag
            Me.Visible = False

        Case vbKeyEscape
            m_Index = -1
            Me.Visible = False
            
    End Select
    
End Sub

Private Sub LpChangeMonth_Click(Index As Integer)
  Select Case LpMonthYear.Caption
    Case StrConv(Format(xDate, "MMMM YYYY"), vbProperCase)
      xDate = IIf(Index = 0, DateAdd("m", -1, xDate), DateAdd("m", 1, xDate))
      ChangeDate
      
    Case Trim(Year(xDate))
      xDate = IIf(Index = 0, DateAdd("yyyy", -1, xDate), DateAdd("yyyy", 1, xDate))
      ChangeMonth
      
    Case Trim(Year(xDate) - 9) & "-" & Trim(Year(xDate))
      xDate = IIf(Index = 0, DateAdd("yyyy", -10, xDate), DateAdd("yyyy", 10, xDate))
      ChangeYear
  End Select
End Sub

Private Sub LpChangeMonth_MouseEnter(Index As Integer)
  Changed
End Sub

Private Sub LpChangeMonth_MouseLeave(Index As Integer)
  Changed
End Sub

Private Sub LpMonth_Click(Index As Integer)
  xDate = CDate(str(Day(xDate)) & "/" & LpMonth(Index).Tag & "/" & str(Year(xDate)))
  ChangeDate
End Sub

Private Sub LpMonth_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = Index
  Changed
End Sub

Private Sub LpMonth_MouseEnter(Index As Integer)
  Changed
End Sub

Private Sub LpMonth_MouseLeave(Index As Integer)
  Changed
End Sub

Private Sub LpMonth_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = -1
  Changed
End Sub

Private Sub LpMonthYear_Click()
  fraDay.Visible = False
  
  Select Case LpMonthYear.Caption
    Case StrConv(Format(xDate, "MMMM YYYY"), vbProperCase)
      LpMonthYear.Caption = Trim(Year(xDate))
      ChangeMonth
      
    Case Trim(Year(xDate))
      LpMonthYear.Caption = Trim(Year(xDate) - 9) & "-" & Trim(Year(xDate))
      ChangeYear
  End Select
End Sub

Private Sub LpMonthYear_MouseEnter()
  Changed
End Sub

Private Sub LpMonthYear_MouseLeave()
  Changed
End Sub

Private Sub LpNum_Click(Index As Integer)
    m_Index = Index
    DateValue = LpNum(Index).Tag
    Me.Visible = False
End Sub

Private Sub LpNum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = Index
  Changed
End Sub

Private Sub LpNum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = -1
  Changed
End Sub

Private Sub LpNum_MouseEnter(Index As Integer)
  Changed
End Sub

Private Sub LpNum_MouseLeave(Index As Integer)
  Changed
End Sub

Private Sub LpToday_Click()
  m_Index = 0
  DateValue = Date
  ChangeDate
  Me.Visible = False
End Sub

Private Sub LpToday_MouseEnter()
  Changed
End Sub

Private Sub LpToday_MouseLeave()
  Changed
End Sub

Private Sub LpYear_Click(Index As Integer)
  xDate = CDate(str(Day(xDate)) & "/" & str(Month(xDate)) & "/" & LpYear(Index).Tag)
  ChangeMonth
End Sub

Private Sub LpYear_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = Index
  Changed
End Sub

Private Sub LpYear_MouseEnter(Index As Integer)
  Changed
End Sub

Private Sub LpYear_MouseLeave(Index As Integer)
  Changed
End Sub

Private Sub LpYear_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_Index = -1
  Changed
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cCLW.DestroyCanvas
  Set cCLW = Nothing
End Sub

'// Private Functions and Sub
Private Function FirstDayOfWeek(xDate As Date) As Date
  Dim d As Integer
  
  d = Weekday(xDate, vbUseSystemDayOfWeek)
  FirstDayOfWeek = DateAdd("d", -d + 1, xDate)
End Function

Private Function ChangeDate()
  Dim i As Long
  Dim d As Integer
  Dim FDOM As Date
  Dim Days As Date
  
  
  fraDay.Visible = True
  fraMonth.Visible = False
  fraYear.Visible = False
  
  
  FDOM = DateSerial(Year(xDate), Month(xDate), 1)
  
  d = Weekday(FDOM, vbUseSystemDayOfWeek)
  For i = 1 To LpNum.Count
    With LpNum(i - 1)
      Days = DateAdd("d", -d + i, FDOM)
      
      .Caption = Day(Days)
      .Tag = Days
      
      If Month(FDOM) = Month(Days) Then
        If Days = Date Then
          .Border = True
          .BorderColor = ForeColorSelectMonth
        Else
          .Border = False
        End If
        
        If Days = FormatDateTime(DateValue, vbShortDate) Then
          .BackColor = ForeColorSelectMonth
          .ForeColor = vbWhite
          If m_Index = -1 Then m_Index = i - 1
        Else
          .BackColor = vbBlack
          .ForeColor = ForeColorSelectMonth
        End If
      Else
        .ForeColor = ForeColorOtherMonth
      End If
    End With
  Next
  
  LpMonthYear.Caption = StrConv(Format(xDate, "MMMM YYYY"), vbProperCase)

  Call Changed
End Function

Private Function ChangeMonth()
  Dim i As Long
  Dim FDOM As Date
  

  fraDay.Visible = False
  fraMonth.Visible = True
  fraYear.Visible = False
  
  LpMonthYear.Caption = Trim(Year(xDate))
  
  FDOM = DateSerial(Year(xDate), Month(xDate), 1)
  
  For i = 0 To LpMonth.Count - 1
    With LpMonth(i)
      .Caption = StrConv(MonthName(i + 1, True), vbProperCase)
      .Tag = i + 1
      
      If (i + 1) = Month(FDOM) And LpMonthYear.Caption = Trim(Year(Now)) Then
        .BorderColor = ForeColorSelectMonth
        .Border = True
        .BackColor = ForeColorSelectMonth
        .BackColorOpacity = 100
        .ForeColor = vbWhite
      Else
        .Border = False
        .BackColor = vbBlack
        .ForeColor = ForeColorSelectMonth
      End If
    End With
  Next
  
  Call Changed
End Function

Private Function ChangeYear()
  Dim i As Long
  Dim FDOM As Date
  Dim lngYearStart As Long
  Dim lngYearEnd As Long
  
  fraDay.Visible = False
  fraMonth.Visible = False
  fraYear.Visible = True
  
  lngYearStart = Year(xDate) - 10
  lngYearEnd = Year(xDate) + 1
  
  LpMonthYear.Caption = Trim(Year(xDate) - 9) & "-" & Trim(Year(xDate))
  
  FDOM = DateSerial(Year(xDate), Month(xDate), 1)
  
  For i = 0 To LpYear.Count - 1
    With LpYear(i)
      .Caption = str(lngYearStart + i)
      .Tag = str(lngYearStart + i)
      
      If lngYearStart + i = Trim(Year(Now)) Then
        .BorderColor = ForeColorSelectMonth
        .Border = True
        .BackColor = ForeColorSelectMonth
        .BackColorOpacity = 100
        .ForeColor = vbWhite
      Else
        .Border = False
        .BackColor = vbBlack
        .ForeColor = ForeColorSelectMonth
      End If
    End With
  Next
  
  Call Changed
End Function

Private Sub Update()
  Dim i As Long
  Dim LP As LabelPlus
  
' cCLW.Clear
  
  For i = Me.Controls.Count - 1 To 0 Step -1
    If TypeName(Me.Controls(i)) = "LabelPlus" Then
        With Me.Controls(i)
          Select Case .Name
            Case "LpNum"
        
              If .IsMouseEnter Then
                If .BackColor = vbBlack Then
                  If .Index = m_Index Then
                   
                    .BackColorOpacity = 15
                  Else
                    .BackColorOpacity = 5
                  End If
                Else
                  .BackColorOpacity = 80
                End If
              Else
                If .BackColor = vbBlack Then
                  If .Index = m_Index Then
                    .BackColorOpacity = 5
                  Else
                    .BackColorOpacity = 0
                  End If
                Else
                  .BackColorOpacity = 100
                End If
              End If
              
            Case "LpMonth", "LpYear"
              If .IsMouseEnter Then
                If .BackColor = vbBlack Then
                  If .Index = m_Index Then
                    .BackColorOpacity = 15
                  Else
                    .BackColorOpacity = 5
                  End If
                Else
                  .BackColorOpacity = 80
                End If
              Else
                If .BackColor = vbBlack Then
                  .BackColorOpacity = 0
                Else
                  .BackColorOpacity = 100
                End If
              End If
            
            Case "LpToday"
              If .IsMouseEnter Then
                .ForeColor = ResalteColor
                .BorderColor = ResalteColor
              Else
                .ForeColor = ThemeColor
                .BorderColor = ThemeColor
              End If
              
            Case "LpMonthYear"
              If .IsMouseEnter Then
                .ForeColor = ThemeColor
                .BorderColor = ThemeColor
                .Border = True
                .BackColorOpacity = 100
              Else
                .ForeColor = ThemeColor
                .BorderColor = &HEEEEEE
                .Border = False
                .BackColorOpacity = 0
              End If
          End Select
    
        If TypeName(.Container) <> "Frame" Then
            .Draw cCLW.hdc, 0, .Left, .Top
        Else
            If .Container.Visible Then
                .Draw cCLW.hdc, 0, .Container.Left + (.Left / Screen.TwipsPerPixelX), .Container.Top + (.Top / Screen.TwipsPerPixelY)
            End If
        End If
        
        End With
    End If
  Next
  
  cCLW.UpdateLayered hwnd
  
End Sub

Private Sub Changed()
    TimerUpdate.Interval = 1
End Sub

Private Sub TimerUpdate_Timer()
    TimerUpdate.Interval = 0
    Update
End Sub
