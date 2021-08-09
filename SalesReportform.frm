VERSION 5.00
Begin VB.Form SalesReportForm 
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
      Begin VB.Image lblDelete 
         Height          =   495
         Left            =   18240
         Picture         =   "SalesReportform.frx":0000
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image lblDivide 
         Height          =   375
         Left            =   19530
         Picture         =   "SalesReportform.frx":17DD
         Stretch         =   -1  'True
         Top             =   6225
         Width           =   375
      End
      Begin ShopManagementSystem.LabelPlus lblOnOff 
         Height          =   735
         Left            =   15720
         TabIndex        =   43
         Top             =   7080
         Width           =   735
         _extentx        =   1296
         _extenty        =   1296
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":1F4A
         font            =   "SalesReportform.frx":1F76
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":1FAA
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblMultiply 
         Height          =   735
         Left            =   19200
         TabIndex        =   42
         Top             =   4920
         Width           =   975
         _extentx        =   1720
         _extenty        =   1296
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":1FD6
         font            =   "SalesReportform.frx":1FF8
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":202C
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblPercentage 
         Height          =   855
         Left            =   16920
         TabIndex        =   41
         Top             =   2760
         Width           =   735
         _extentx        =   1296
         _extenty        =   1508
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2058
         font            =   "SalesReportform.frx":207A
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":20AE
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   495
         Index           =   10
         Left            =   18240
         TabIndex        =   40
         Top             =   7080
         Width           =   495
         _extentx        =   873
         _extenty        =   873
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":20DA
         font            =   "SalesReportform.frx":20FC
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2130
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblEqual 
         Height          =   855
         Left            =   19320
         TabIndex        =   39
         Top             =   7080
         Width           =   735
         _extentx        =   1296
         _extenty        =   1508
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":215C
         font            =   "SalesReportform.frx":217E
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":21B2
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblAdd 
         Height          =   855
         Left            =   19320
         TabIndex        =   38
         Top             =   2760
         Width           =   735
         _extentx        =   1296
         _extenty        =   1508
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":21DE
         font            =   "SalesReportform.frx":2200
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2234
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblSubtract 
         Height          =   615
         Left            =   19390
         TabIndex        =   37
         Top             =   3940
         Width           =   615
         _extentx        =   1085
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2260
         font            =   "SalesReportform.frx":2282
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":22B6
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   0
         Left            =   15880
         TabIndex        =   36
         Top             =   3960
         Width           =   375
         _extentx        =   661
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":22E2
         font            =   "SalesReportform.frx":2304
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2338
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   735
         Index           =   3
         Left            =   15840
         TabIndex        =   35
         Top             =   4980
         Width           =   495
         _extentx        =   873
         _extenty        =   1296
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2364
         font            =   "SalesReportform.frx":2386
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":23BA
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   6
         Left            =   15840
         TabIndex        =   34
         Top             =   6120
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":23E6
         font            =   "SalesReportform.frx":2408
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":243C
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   735
         Index           =   1
         Left            =   17040
         TabIndex        =   33
         Top             =   3880
         Width           =   495
         _extentx        =   873
         _extenty        =   1296
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2468
         font            =   "SalesReportform.frx":248A
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":24BE
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   735
         Index           =   4
         Left            =   17040
         TabIndex        =   32
         Top             =   4980
         Width           =   495
         _extentx        =   873
         _extenty        =   1296
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":24EA
         font            =   "SalesReportform.frx":250C
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2540
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   2
         Left            =   18240
         TabIndex        =   31
         Top             =   3960
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":256C
         font            =   "SalesReportform.frx":258E
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":25C2
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   5
         Left            =   18240
         TabIndex        =   30
         Top             =   5040
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":25EE
         font            =   "SalesReportform.frx":2610
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2644
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   7
         Left            =   17040
         TabIndex        =   29
         Top             =   6120
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2670
         font            =   "SalesReportform.frx":2692
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":26C6
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   8
         Left            =   18240
         TabIndex        =   28
         Top             =   6120
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":26F2
         font            =   "SalesReportform.frx":2714
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":2748
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lbl 
         Height          =   615
         Index           =   9
         Left            =   17040
         TabIndex        =   27
         Top             =   7200
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2774
         font            =   "SalesReportform.frx":2796
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":27CA
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.LabelPlus lblC 
         Height          =   615
         Left            =   15840
         TabIndex        =   26
         Top             =   2880
         Width           =   495
         _extentx        =   873
         _extenty        =   1085
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   1
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":27F6
         font            =   "SalesReportform.frx":2818
         forecolor       =   14737632
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":284C
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnOnOff 
         Height          =   1215
         Left            =   15480
         TabIndex        =   25
         Top             =   6840
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   9
         Left            =   16680
         TabIndex        =   24
         Top             =   6840
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDot 
         Height          =   1215
         Left            =   17880
         TabIndex        =   23
         Top             =   6840
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnEqual 
         Height          =   1215
         Left            =   19080
         TabIndex        =   22
         Top             =   6840
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   6
         Left            =   15480
         TabIndex        =   21
         Top             =   5760
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   7
         Left            =   16680
         TabIndex        =   20
         Top             =   5760
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   8
         Left            =   17880
         TabIndex        =   19
         Top             =   5760
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDivide 
         Height          =   1215
         Left            =   19080
         TabIndex        =   18
         Top             =   5760
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   3
         Left            =   15480
         TabIndex        =   17
         Top             =   4680
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   4
         Left            =   16680
         TabIndex        =   16
         Top             =   4680
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   5
         Left            =   17880
         TabIndex        =   15
         Top             =   4680
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnMultiply 
         Height          =   1215
         Left            =   19080
         TabIndex        =   14
         Top             =   4680
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   0
         Left            =   15480
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   1
         Left            =   16680
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btn 
         Height          =   1215
         Index           =   2
         Left            =   17880
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnSubtract 
         Height          =   1215
         Left            =   19080
         TabIndex        =   10
         Top             =   3600
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnDelete 
         Height          =   1215
         Left            =   17880
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnAdd 
         Height          =   1215
         Left            =   19080
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnPercentage 
         Height          =   1215
         Left            =   16680
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism btnC 
         Height          =   1215
         Left            =   15480
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.LabelPlus lblDisplay 
         Height          =   735
         Left            =   15600
         TabIndex        =   5
         Top             =   960
         Width           =   4695
         _extentx        =   8281
         _extenty        =   1296
         backcolor       =   4210752
         backcoloropacity=   0
         backshadow      =   0   'False
         captionalignmenth=   2
         captionalignmentv=   1
         caption         =   "SalesReportform.frx":2878
         font            =   "SalesReportform.frx":2898
         forecolor       =   8421504
         shadowcoloropacity=   0
         calloutalign    =   0
         calloutwidth    =   0
         calloutlen      =   0
         mousepointer    =   0
         iconfont        =   "SalesReportform.frx":28C0
         iconforecolor   =   0
         iconopacity     =   0
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
         Height          =   1215
         Left            =   15240
         TabIndex        =   4
         Top             =   720
         Width           =   5295
         _extentx        =   9340
         _extenty        =   2143
         distance        =   5
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin ShopManagementSystem.ucNeumorphism calculator 
         Height          =   8775
         Left            =   15000
         TabIndex        =   3
         Top             =   120
         Width           =   5775
         _extentx        =   10186
         _extenty        =   15478
         blur            =   5
         backcolor       =   4210752
         mousepointer    =   0
      End
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   1335
      Left            =   17280
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
      _extentx        =   2355
      _extenty        =   2355
      distance        =   2
      blur            =   5
      mousepointer    =   0
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
      Picture         =   "SalesReportform.frx":28EC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20775
   End
End
Attribute VB_Name = "SalesReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit
Dim count1 As Integer
Dim result As Integer



Private Sub btn_Click(Index As Integer)
If count1 = 0 Then
lblDisplay.Caption = " "
MsgBox ("Calculator is not on")

End If

If count1 = 1 Then
lblDisplay.Caption = " "
count1 = count1 + 1
End If

If count1 > 1 Then
lblDisplay.Caption = lblDisplay.Caption & lbl(Index).Caption
End If
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

Private Sub Form_Load()


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
lblDisplay.Caption = lblDisplay.Caption & lbl(Index).Caption
End If
End Sub

Private Sub lblOnOff_Click()
Call btnOnOff_Click
End Sub
