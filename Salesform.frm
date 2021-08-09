VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SalesForm 
   BorderStyle     =   0  'None
   Caption         =   "Sales"
   ClientHeight    =   10395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20805
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   20805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   20775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7560
         TabIndex        =   51
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118489089
         CurrentDate     =   44316
      End
      Begin MSDataGridLib.DataGrid DailyDataGrid 
         Height          =   6135
         Left            =   915
         TabIndex        =   49
         Top             =   1680
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   10821
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin ShopManagementSystem.ucNeumorphism DailyDataBackground 
         Height          =   6615
         Left            =   720
         TabIndex        =   48
         Top             =   1440
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   11668
         Distance        =   5
         Blur            =   3
         LightColor      =   8421504
         MousePointer    =   0
      End
      Begin ShopManagementSystem.LabelPlus lblData 
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   46
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   4210752
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         HotLineColor    =   8421504
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
      Begin ShopManagementSystem.LabelPlus lblData 
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   45
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   4210752
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":003C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         HotLine         =   -1  'True
         HotLineColor    =   8421504
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
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism4 
         Height          =   8295
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   14631
         Distance        =   5
         LightColor      =   0
         BackColor       =   4210752
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
         Height          =   8895
         Left            =   -120
         TabIndex        =   44
         Top             =   120
         Width           =   15375
         _ExtentX        =   27120
         _ExtentY        =   15690
         Blur            =   10
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.LabelPlus lblDisplay 
         Height          =   735
         Left            =   15600
         TabIndex        =   43
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1296
         BackColor       =   4210752
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   2
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":006C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632256
         ShadowColorOpacity=   0
         CallOutAlign    =   0
         CallOutWidth    =   0
         CallOutLen      =   0
         MousePointer    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   12632256
         IconOpacity     =   0
         PictureArr      =   0
      End
      Begin VB.Image lblDelete 
         Height          =   495
         Left            =   18240
         Picture         =   "Salesform.frx":008C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image lblDivide 
         Height          =   375
         Left            =   19530
         Picture         =   "Salesform.frx":1869
         Stretch         =   -1  'True
         Top             =   6225
         Width           =   375
      End
      Begin ShopManagementSystem.LabelPlus lblOnOff 
         Height          =   735
         Left            =   15720
         TabIndex        =   42
         Top             =   7080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":1FD6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblMultiply 
         Height          =   735
         Left            =   19200
         TabIndex        =   41
         Top             =   4920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2002
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblPercentage 
         Height          =   855
         Left            =   16920
         TabIndex        =   40
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1508
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   495
         Index           =   10
         Left            =   18240
         TabIndex        =   39
         Top             =   7080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2046
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblEqual 
         Height          =   855
         Left            =   19320
         TabIndex        =   38
         Top             =   7080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1508
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2068
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblAdd 
         Height          =   855
         Left            =   19320
         TabIndex        =   37
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1508
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":208A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblSubtract 
         Height          =   615
         Left            =   19390
         TabIndex        =   36
         Top             =   3940
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":20AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   0
         Left            =   15880
         TabIndex        =   35
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":20CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   735
         Index           =   3
         Left            =   15840
         TabIndex        =   34
         Top             =   4980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1296
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":20F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   6
         Left            =   15840
         TabIndex        =   33
         Top             =   6120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2112
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   735
         Index           =   1
         Left            =   17040
         TabIndex        =   32
         Top             =   3880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1296
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2134
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   735
         Index           =   4
         Left            =   17040
         TabIndex        =   31
         Top             =   4980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1296
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2156
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   2
         Left            =   18240
         TabIndex        =   30
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2178
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   5
         Left            =   18240
         TabIndex        =   29
         Top             =   5040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":219A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   7
         Left            =   17040
         TabIndex        =   28
         Top             =   6120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":21BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   8
         Left            =   18240
         TabIndex        =   27
         Top             =   6120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":21DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus Lbl 
         Height          =   615
         Index           =   9
         Left            =   17040
         TabIndex        =   26
         Top             =   7200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2200
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.LabelPlus lblC 
         Height          =   615
         Left            =   15840
         TabIndex        =   25
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         BackColorOpacity=   0
         BackShadow      =   0   'False
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "Salesform.frx":2222
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14737632
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
      Begin ShopManagementSystem.ucNeumorphism btnOnOff 
         Height          =   1215
         Left            =   15480
         TabIndex        =   24
         Top             =   6840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   9
         Left            =   16680
         TabIndex        =   23
         Top             =   6840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDot 
         Height          =   1215
         Left            =   17880
         TabIndex        =   22
         Top             =   6840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnEqual 
         Height          =   1215
         Left            =   19080
         TabIndex        =   21
         Top             =   6840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   6
         Left            =   15480
         TabIndex        =   20
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   7
         Left            =   16680
         TabIndex        =   19
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   8
         Left            =   17880
         TabIndex        =   18
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDivide 
         Height          =   1215
         Left            =   19080
         TabIndex        =   17
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   3
         Left            =   15480
         TabIndex        =   16
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   4
         Left            =   16680
         TabIndex        =   15
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   5
         Left            =   17880
         TabIndex        =   14
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnMultiply 
         Height          =   1215
         Left            =   19080
         TabIndex        =   13
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   0
         Left            =   15480
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   1
         Left            =   16680
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   2
         Left            =   17880
         TabIndex        =   10
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnSubtract 
         Height          =   1215
         Left            =   19080
         TabIndex        =   9
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDelete 
         Height          =   1215
         Left            =   17880
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnAdd 
         Height          =   1215
         Left            =   19080
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnPercentage 
         Height          =   1215
         Left            =   16680
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnC 
         Height          =   1215
         Left            =   15480
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
         Height          =   1215
         Left            =   15240
         TabIndex        =   4
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2143
         Distance        =   5
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism calculator 
         Height          =   8775
         Left            =   15000
         TabIndex        =   3
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   15478
         Blur            =   5
         BackColor       =   4210752
         MousePointer    =   0
      End
   End
   Begin ShopManagementSystem.LabelPlus lblData 
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   50
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   4210752
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Salesform.frx":2244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      HotLineColor    =   8421504
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
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   1335
      Left            =   17280
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      Distance        =   2
      Blur            =   5
      MousePointer    =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.Image Sales_Background_Image 
      Height          =   10335
      Left            =   0
      Picture         =   "Salesform.frx":2280
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20775
   End
End
Attribute VB_Name = "SalesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Dim op As String
Option Explicit

Dim i As Integer
Dim exp1 As Double
Dim exp2 As Double
Dim result As Double
Dim count1 As Integer
Dim sign As String
Public VarDailyLogSheet As ADODB.Recordset





Private Sub lblDaily_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Long
    lblDaily.ZOrder
    For i = 0 To lblDaily.Count - 1
        lblDaily.HotLine = True
        lblDaily.ForeColor = IIf(vbWhite, &HDED0BC)
    Next
End Sub

Private Sub lblDaily_MouseEnter()
    lblDaily.ForeColor = vbWhite
End Sub

Private Sub lblDaily_MouseLeave()
    If lblDaily.HotLine = False Then
        lblDaily.ForeColor = &HDED0BC
    End If
End Sub

Private Sub btn_Click(Index As Integer)
If count1 = 0 Then
lblDisplay.Caption = " "
MsgBox ("Calculator is not on")

End If




'VarLoginForm.Find "UserName='" & usertxt1 & "'", 0, adSearchForward, start

If count1 = 1 Then
lblDisplay.Caption = " "
count1 = count1 + 1
End If

If count1 > 1 Then
lblDisplay.Caption = lblDisplay.Caption & Lbl(Index).Caption
End If
End Sub










Private Sub btnC_Click()
result = 0
exp1 = 0
exp2 = 0
lblDisplay.Caption = " "
count1 = 1
End Sub



Private Sub btnEqual_Click()
exp2 = Val(lblDisplay.Caption)


End Sub

Private Sub btnOnOff_Click()
If lblDisplay.Caption = "" Then
result = 0
count1 = 1
lblDisplay.Caption = "0"
Else
count1 = 0
lblDisplay.Caption = ""
End If
End Sub

Private Sub btnOp_Click(Index As Integer)
result = exp1
exp1 = result + Val(lblDisplay.Caption)
lblDisplay.Caption = " "
op = btnOp(Index).Caption
End Sub



Private Sub Command1_Click()

        LogSheetReport.BottomMargin = 0
        LogSheetReport.LeftMargin = 0
        LogSheetReport.RightMargin = 0
        LogSheetReport.TopMargin = 0
        '------ report.refresh
        'report.Show
LogSheetReport.Show
End Sub

Private Sub Form_Load()
DTPicker1.Visible = False
LogSheetReport.Visible = False


ucNeumorphism2.Visible = False

Sales_Background_Image.Move 0, 0, Me.Width, Me.Height

initialcontrollist = GetLocation(Me)
ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top


End Sub

Private Sub Form_Resize()
    Sales_Background_Image.Width = Me.ScaleWidth
    Sales_Background_Image.Height = Me.ScaleHeight
    
    ResizeControls Me, initialcontrollist
End Sub





Private Sub lbl_Click(Index As Integer)
If count1 = 0 Then
lblDisplay.Caption = " "
MsgBox ("Calculator is not on")

End If

If count1 = 1 Then
lblDisplay.Caption = " "
count1 = count1 + 1
End If

If count1 > 1 Then
lblDisplay.Caption = lblDisplay.Caption & Lbl(Index).Caption
End If
End Sub

Private Sub lblAdd_Click()
exp1 = lblDisplay.Caption
sign = "+"
lblDisplay.Caption = ""
End Sub

Private Sub lblC_Click()
result = 0
exp1 = 0
exp2 = 0
lblDisplay.Caption = " "
count1 = 1
End Sub



Private Sub lblDaily_Click()

End Sub

Private Sub lblData_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Long
    lblData(Index).ZOrder
    For i = 0 To lblData.Count - 1
        lblData(i).HotLine = CBool(i = Index)
        
        lblData(i).ForeColor = IIf(i = Index, vbWhite, &HDED0BC)
        
    Next
End Sub

Private Sub lblData_MouseEnter(Index As Integer)
    
    
    
If Index = 1 Then
    DailyDataGrid.Visible = True
    DTPicker1.Visible = True
    
    
Prev.Open "select * from temp1", cn, adOpenDynamic, adLockOptimistic
Dateprev.Open "select * from LogSheet where Date_of_Printing = # " & DTPicker1.Value & " # ", cn, adOpenDynamic, adLockOptimistic


On Error GoTo BB

Dateprev.MoveFirst
While (Dateprev.EOF = False)

    Prev.AddNew
    Prev.Fields("Date_of_Printing") = Dateprev.Fields("DATE_OF_PRINTING")
    Prev.Fields("Job_Name") = Dateprev.Fields("JOB_NAME")
    Prev.Fields("Client") = Dateprev.Fields("CLIENT")
    Prev.Fields("Quantity") = Dateprev.Fields("QUANTITY")
    Prev.Fields("Price") = Dateprev.Fields("PRICE")
    Prev.Fields("Total_Cost") = Dateprev.Fields("TOTAL_COST")
    Dateprev.MoveNext
    Prev.Update

Wend
Prev.Close
Dateprev.Close

        LogSheetReport.BottomMargin = 0
        LogSheetReport.LeftMargin = 0
        LogSheetReport.RightMargin = 0
        LogSheetReport.TopMargin = 0
LogSheetReport.Show
cn.Execute ("delete * from temp1")

Exit Sub

BB:
    MsgBox "Select Another Date!", vbCritical
    Prev.Close
    Dateprev.Close


ElseIf Index = 0 Then
DTPicker1.Visible = False
LogSheetReport.Visible = False
    DailyDataGrid.Visible = True
    Dim VarDailyLogSheet As New Recordset
    VarDailyLogSheet.Open "select * from LogSheet", cn, adOpenDynamic, adLockOptimistic
    Set DailyDataGrid.DataSource = VarDailyLogSheet
    DailyDataGrid.Refresh
    Set VarDailyLogSheet = Nothing
  
    
End If
    lblData(Index).ForeColor = vbWhite
    
End Sub

Private Sub lblData_MouseLeave(Index As Integer)

    If lblData(Index).HotLine = False Then
        lblData(Index).ForeColor = &HDED0BC
        
    End If
    
    
End Sub

Private Sub lblDelete_Click()


If lblDisplay.Caption = " " Then
result = 0
exp1 = 0
exp2 = 0
lblDisplay.Caption = " "
count1 = 1
Else

lblDisplay.Caption = Left(lblDisplay.Caption, Len(lblDisplay.Caption) - 1)

'lblDisplay.SetFocus
'lblDisplay.SelStart = Len(lblDisplay.caption)
'SendKeys ("{BACKSPACE}")

End If
End Sub

Private Sub lblDivide_Click()
exp1 = lblDisplay.Caption
sign = "/"
lblDisplay.Caption = ""
End Sub

Private Sub lblEqual_Click()

exp2 = Val(lblDisplay.Caption)

If sign = "+" Then
        result = exp1 + exp2
        lblDisplay.Caption = result
        count1 = 0
Else
If sign = "-" Then
        result = exp1 - exp2
        lblDisplay.Caption = result
        count1 = 0
Else
If sign = "*" Then
        result = exp1 * exp2
        lblDisplay.Caption = result
        count1 = 0
Else
If sign = "/" Then
        result = exp1 / exp2
        lblDisplay.Caption = result
        count1 = 0
Else
If sign = "%" Then
        result = (exp1 / 100) * exp2
        lblDisplay.Caption = result
        count1 = 0

End If
End If
End If
End If
End If
End Sub

Private Sub lblMultiply_Click()
exp1 = lblDisplay.Caption
sign = "*"
lblDisplay.Caption = ""
End Sub

Private Sub lblOnOff_Click()
Call btnOnOff_Click
End Sub





Private Sub lblPercentage_Click()
exp1 = lblDisplay.Caption
sign = "%"
lblDisplay.Caption = ""
End Sub

Private Sub lblSubtract_Click()
exp1 = lblDisplay.Caption
sign = "-"
lblDisplay.Caption = ""
End Sub
