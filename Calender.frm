VERSION 5.00
Begin VB.Form Calender 
   BorderStyle     =   0  'None
   Caption         =   "Calender"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerUpdate 
      Left            =   2640
      Top             =   3840
   End
   Begin ShopManagementSystem.LabelPlus LpToday 
      Height          =   255
      Left            =   360
      TabIndex        =   46
      Top             =   3720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColorOpacity=   50
      BackShadow      =   0   'False
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      Caption         =   "Calender.frx":0000
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
   Begin ShopManagementSystem.LabelPlus LabelPlus4 
      Height          =   30
      Index           =   1
      Left            =   360
      TabIndex        =   45
      Top             =   3960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   53
      BackColor       =   12632256
      BackShadow      =   0   'False
      BorderColor     =   4210752
      Caption         =   "Calender.frx":002A
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
   Begin ShopManagementSystem.LabelPlus LabelPlus4 
      Height          =   30
      Index           =   0
      Left            =   360
      TabIndex        =   44
      Top             =   3720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   53
      BackColor       =   12632256
      BackShadow      =   0   'False
      BorderColor     =   4210752
      Caption         =   "Calender.frx":005E
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   34
      Left            =   2520
      TabIndex        =   43
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0092
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   33
      Left            =   2160
      TabIndex        =   42
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":00B4
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   32
      Left            =   1800
      TabIndex        =   41
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":00D6
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   31
      Left            =   1440
      TabIndex        =   40
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":00F8
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   30
      Left            =   1080
      TabIndex        =   39
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":011A
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   29
      Left            =   720
      TabIndex        =   38
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":013C
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   28
      Left            =   360
      TabIndex        =   37
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":015E
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   27
      Left            =   2520
      TabIndex        =   36
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0180
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   26
      Left            =   2160
      TabIndex        =   35
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":01A2
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   25
      Left            =   1800
      TabIndex        =   34
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":01C4
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   24
      Left            =   1440
      TabIndex        =   33
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":01E6
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   23
      Left            =   1080
      TabIndex        =   32
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0208
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   22
      Left            =   720
      TabIndex        =   31
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":022A
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   21
      Left            =   360
      TabIndex        =   30
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":024C
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   20
      Left            =   2520
      TabIndex        =   29
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":026E
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   19
      Left            =   2160
      TabIndex        =   28
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0290
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   18
      Left            =   1800
      TabIndex        =   27
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":02B2
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   17
      Left            =   1440
      TabIndex        =   26
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":02D4
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   16
      Left            =   1080
      TabIndex        =   25
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":02F6
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   15
      Left            =   720
      TabIndex        =   24
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0318
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   14
      Left            =   360
      TabIndex        =   23
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":033A
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   13
      Left            =   2520
      TabIndex        =   22
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":035C
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   12
      Left            =   2160
      TabIndex        =   21
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":037E
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   11
      Left            =   1800
      TabIndex        =   20
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":03A0
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   10
      Left            =   1440
      TabIndex        =   19
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":03C2
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   18
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":03E4
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   8
      Left            =   720
      TabIndex        =   17
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0406
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0428
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":044A
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":046C
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   13
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":048E
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   12
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":04B0
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":04D2
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
   Begin ShopManagementSystem.LabelPlus LabelPlus3 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":04F4
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
   Begin ShopManagementSystem.LabelPlus LpNum 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0516
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":0538
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   5
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":055C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":0580
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   10
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":05A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":05C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":05EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LabelPlus2 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      Caption         =   "Calender.frx":0610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin ShopManagementSystem.LabelPlus LpMonthYear 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackShadow      =   0   'False
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Calender.frx":0634
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
   Begin VB.Image ImgChangeMonth 
      Height          =   255
      Index           =   0
      Left            =   360
      Picture         =   "Calender.frx":0668
      Stretch         =   -1  'True
      Top             =   650
      Width           =   135
   End
   Begin VB.Image LpChangeMonth 
      Height          =   255
      Index           =   1
      Left            =   2760
      Picture         =   "Calender.frx":1075
      Stretch         =   -1  'True
      Top             =   650
      Width           =   135
   End
   Begin ShopManagementSystem.LabelPlus lblCalender 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3060
      _ExtentX        =   6800
      _ExtentY        =   8070
      BackColor       =   16777215
      Border          =   -1  'True
      BorderColorOpacity=   20
      BorderCornerLeftTop=   4
      BorderCornerRightTop=   4
      BorderCornerBottomRight=   4
      BorderCornerBottomLeft=   4
      BorderWidth     =   1
      Caption         =   "Calender.frx":1A82
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
      ShadowColorOpacity=   0
      CallOutPosicion =   1
      CallOutWidth    =   20
      CallOut         =   -1  'True
      CallOutCustomPosition=   180
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
Attribute VB_Name = "Calender"
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
  
  cCLW.Clear
  
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

