VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DailyLogSheet 
   BorderStyle     =   0  'None
   ClientHeight    =   8715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   16305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalculate 
      Appearance      =   0  'Flat
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Frame DailyLogSheetFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Daily Printing System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7095
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   16335
      Begin VB.Frame ReceiptFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Receipt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   6735
         Left            =   10800
         TabIndex        =   2
         Top             =   240
         Width           =   5415
         Begin VB.Image Signaturewhite 
            Height          =   1455
            Left            =   3120
            Picture         =   "DailyLogSheet.frx":0000
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   2295
         End
         Begin VB.Label TotalCostlbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   49
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Pricelbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   48
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Quantitylbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   47
            Top             =   3600
            Width           =   1935
         End
         Begin VB.Label ProductNamelbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   46
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label CustomerNamelbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   45
            Top             =   2325
            Width           =   1935
         End
         Begin VB.Label Signature1lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Signature"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3480
            TabIndex        =   44
            Top             =   6360
            Width           =   1815
         End
         Begin VB.Label Datelbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   6000
            Width           =   1695
         End
         Begin VB.Label DTlbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DATE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   6360
            Width           =   1695
         End
         Begin VB.Label TClbl 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL COST"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Plbl 
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Qtylbl 
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   3600
            Width           =   1935
         End
         Begin VB.Label Pnamelbl 
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT NAME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Cnamelbl 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER NAME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   2325
            Width           =   1935
         End
         Begin VB.Label line1lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   0
            TabIndex        =   31
            Top             =   1800
            Width           =   5415
         End
         Begin VB.Label ShopAddresslbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SHOP NO 1,SILVASSA,VAPI ROAD (396230)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1680
            TabIndex        =   30
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label ShopContactlbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "9724224417,7874352707"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1680
            TabIndex        =   29
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label ShopNamelbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "S.K PRINTERS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1680
            TabIndex        =   28
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Line2lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   36
            Top             =   2430
            Width           =   5415
         End
         Begin VB.Label Line3lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   37
            Top             =   3100
            Width           =   5415
         End
         Begin VB.Label Line4lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   38
            Top             =   3700
            Width           =   5415
         End
         Begin VB.Label Line5lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   40
            Top             =   4300
            Width           =   5415
         End
         Begin VB.Label Line6lbl 
            BackStyle       =   0  'Transparent
            Caption         =   " ----------------------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   41
            Top             =   4900
            Width           =   5415
         End
         Begin VB.Image Signature2 
            Height          =   1455
            Left            =   3120
            Picture         =   "DailyLogSheet.frx":0F94
            Stretch         =   -1  'True
            Top             =   5280
            Width           =   2295
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9960
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Pricetxt 
         Height          =   315
         Left            =   7680
         TabIndex        =   52
         Top             =   1200
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DailylogsheetDate 
         Height          =   315
         Left            =   4560
         TabIndex        =   27
         Top             =   5040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   16777215
         CalendarTitleForeColor=   4210752
         Format          =   119078913
         CurrentDate     =   44129
      End
      Begin VB.ComboBox cmbColorStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Text            =   "Select Your Choice"
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ComboBox cmbMachine 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4560
         TabIndex        =   25
         Text            =   "Select Your Choice"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.ComboBox cmbPapertype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4560
         TabIndex        =   24
         Text            =   "Select Your Choice"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox cmbPrintStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         ItemData        =   "DailyLogSheet.frx":2046
         Left            =   4560
         List            =   "DailyLogSheet.frx":2048
         TabIndex        =   23
         Text            =   "Select Your Choice"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox qtytxt 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox PaperWeighttxt 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox Clienttxt 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox Sizetxt 
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox JobNametxt 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
         Height          =   6615
         Left            =   10680
         TabIndex        =   54
         Top             =   360
         Width           =   5655
         _extentx        =   9975
         _extenty        =   11668
         mousepointer    =   0
      End
      Begin VB.Label TotalCostlbl2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   240
         TabIndex        =   53
         Top             =   6120
         Width           =   3975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   51
         Top             =   960
         Width           =   1575
      End
      Begin VB.Image Puppy 
         Height          =   3375
         Left            =   7200
         Picture         =   "DailyLogSheet.frx":204A
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Color Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Weight (GSM)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
         Height          =   6015
         Left            =   -480
         TabIndex        =   55
         Top             =   240
         Width           =   11055
         _extentx        =   19500
         _extenty        =   10610
         backcolor       =   4210752
         mousepointer    =   0
      End
   End
   Begin VB.CommandButton cmdPrintReceipt 
      Caption         =   "Print Receipt"
      Height          =   495
      Left            =   12000
      TabIndex        =   50
      Top             =   8040
      Width           =   3615
   End
   Begin VB.CommandButton cmdRecieptGenerate 
      Caption         =   "Generate Receipt"
      Height          =   495
      Left            =   12000
      TabIndex        =   3
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAILY LOG SHEET"
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
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   1815
      Left            =   4200
      TabIndex        =   56
      Top             =   -480
      Width           =   7215
      _extentx        =   12726
      _extenty        =   3201
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin VB.Image DailyLogSheet 
      Height          =   8775
      Left            =   -120
      Picture         =   "DailyLogSheet.frx":65D25
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16455
   End
End
Attribute VB_Name = "DailyLogSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit


Private Function Clear()
    JobNametxt = ""
    Sizetxt = ""
    Clienttxt = ""
    PaperWeighttxt = ""
    qtytxt = ""
    Pricetxt.text = ""
    TotalCostlbl2.Caption = ""
    cmbPrintStatus.text = "Select Your Choice"
    cmbPapertype.text = "Select Your Choice"
    cmbMachine.text = "Select Your Choice"
    cmbColorStatus.text = "Select Your Choice"
    
End Function



Private Sub cmdAdd_Click()
    
    If JobNametxt.text = "" Then
        If Sizetxt.text = "" Then
            If Clienttxt.text = "" Then
                If PaperWeighttxt.text = "" Then
                    If qtytxt.text = "" Then
                        If cmbPrintStatus.text = "Select Your Choice" Then
                            If cmbPapertype.text = "Select Your Choice" Then
                                If cmbMachine.text = "Select Your Choice" Then
                                    If cmbColorStatus.text = "Select Your Choice" Then
                                        If Pricetxt.text = "" Then
                                            MsgBox "Please Enter All The Details Properly", vbOKOnly + vbCritical, App.Title
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        If JobNametxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If Sizetxt.text = "" Then
           MsgBox "Please Enter Paper Size", vbOKOnly + vbCritical, App.Title
        End If
        
        If Clienttxt.text = "" Then
            MsgBox "Please Client Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If PaperWeighttxt.text = "" Then
            MsgBox "Please Enter Paper Weight In GSM", vbOKOnly + vbCritical, App.Title
        End If
        
        If qtytxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
        
        Else
          
            VarLogSheet.AddNew
            
                VarLogSheet.Fields(1) = JobNametxt.text
                VarLogSheet.Fields(2) = Sizetxt.text
                VarLogSheet.Fields(3) = Clienttxt.text
                VarLogSheet.Fields(4) = PaperWeighttxt.text
                VarLogSheet.Fields(5) = Val(qtytxt.text)
                
                VarLogSheet.Fields(6) = cmbPrintStatus.text
                VarLogSheet.Fields(7) = cmbPapertype.text
                VarLogSheet.Fields(8) = cmbMachine.text
                VarLogSheet.Fields(9) = cmbColorStatus.text
                VarLogSheet.Fields(10) = DailylogsheetDate
                
                VarLogSheet.Fields(11) = Pricetxt.text
                VarLogSheet.Fields(12) = qtytxt.text * Pricetxt.text
                
                
                VarStockDetails.Fields(10) = VarStockDetails.Fields(9) + Val(qtytxt.text)
                '-------------------------------------------------------------
                Call cmdRecieptGenerate_Click
                Call cmdCalculate_Click
                '-------------------------------------------------------------
                Dim varstocks As New ADODB.Recordset
                
               
                varstocks.Open "select*from StockDetails", cn, adOpenDynamic, adLockOptimistic
                
                

                If varstocks.Fields(2) = JobNametxt.text Then
                If varstocks.EOF Then
                MsgBox "Product Name " & JobNametxt.text & " not found"
                Else
                varstocks.Fields("Stock_In_(Quantity)") = varstocks.Fields("Stock_In_(Quantity)") - Val(qtytxt.text)
                
                varstocks.Fields("Stock_Out") = varstocks.Fields("Stock_Out") + Val(qtytxt.text)
                
                
                varstocks.Fields("Total_Quantity") = varstocks.Fields("Stock_Out") + varstocks.Fields("Stock_In_(Quantity)")
                
                
                 If varstocks.Fields("Stock_In_(Quantity)") <= 0 Then
                MsgBox "You Don't Have Enough Stock"
                End If
                End If
                
                
                VarLogSheet.Update
               
                Clear
                End If
    End If
    
End Sub

Private Sub cmdCalculate_Click()

If qtytxt.text = "" Then
  If Pricetxt.text = "" Then
    MsgBox "please enter Quantity and Price"
    qtytxt.SetFocus
End If
End If

If qtytxt.text = "" Then
    MsgBox "Please Enter Quantity"
    qtytxt.SetFocus
End If

If Pricetxt.text = "" Then
    MsgBox "Please Enter Price"
    Pricetxt.SetFocus
End If

    
    TotalCostlbl2.Caption = "Your Total Cost is = " & (qtytxt.text * Pricetxt.text)
    
End Sub

Private Sub cmdClear_Click()
    Clear
End Sub

Private Sub cmdExit_Click()
    End
End Sub







Private Sub cmdPrintReceipt_Click()
    'printing a receipt
If MsgBox("Print Receipt?", vbYesNo, "Receipt") = vbYes Then


Printer.Print Tab(30);
Printer.Print Tab(40); "S.K PRINTERS"
Printer.Print Tab(40); "9724224417,7874352707"
Printer.Print Tab(40); "SHOP NO 1,SILVASSA"
Printer.Print Tab(40); "VAPI ROAD (396230)"

Printer.Print
Printer.Print Tab(15); "---------------------------------------------------------------------------------------------------------"
Printer.Print Tab(40); "SERVICE CHARGE RECEIPT"
Printer.Print Tab(15); "---------------------------------------------------------------------------------------------------------"
Printer.Print
Printer.Print Tab(15); "Customer Name:"; Spc(3); Me.Clienttxt.text
Printer.Print
Printer.Print Tab(15); "Product Name:"; Spc(5); Me.cmbPapertype
Printer.Print
Printer.Print Tab(15); "Quantity:"; Spc(9); Me.qtytxt.text
Printer.Print
Printer.Print Tab(15); "Price:"; Spc(9); Me.Pricetxt.text
Printer.Print
Printer.Print Tab(15); "Total Cost:"; Spc(9); Me.qtytxt.text * Pricetxt.text
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(15); "Date:"; Spc(3); Me.DailylogsheetDate
Printer.Print
Printer.Print Tab(15); "Balance B/F:"; Spc(1); ".........................................."
Printer.Print
Printer.Print Tab(15); "Agent's Signature:"; Spc(1); Me.Signaturewhite.Visible = True


Printer.Print
Printer.Print Tab(15); "Agent's Signature:"; Spc(1); "..........................................."
Printer.Print


  
    If JobNametxt.text = "" Then
        If Sizetxt.text = "" Then
            If Clienttxt.text = "" Then
                If PaperWeighttxt.text = "" Then
                    If qtytxt.text = "" Then
                        If cmbPrintStatus.text = "Select Your Choice" Then
                            If cmbPapertype.text = "Select Your Choice" Then
                                If cmbMachine.text = "Select Your Choice" Then
                                    If cmbColorStatus.text = "Select Your Choice" Then
                                        If Pricetxt.text = "" Then
                                            MsgBox "Please Enter All The Details Properly", vbOKOnly + vbCritical, App.Title
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        If JobNametxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If Sizetxt.text = "" Then
           MsgBox "Please Enter Paper Size", vbOKOnly + vbCritical, App.Title
        End If
        
        If Clienttxt.text = "" Then
            MsgBox "Please Client Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If PaperWeighttxt.text = "" Then
            MsgBox "Please Enter Paper Weight In GSM", vbOKOnly + vbCritical, App.Title
        End If
        
        If qtytxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
        
        Else
          
            VarLogSheet.AddNew
            
                VarLogSheet.Fields(1) = JobNametxt.text
                VarLogSheet.Fields(2) = Sizetxt.text
                VarLogSheet.Fields(3) = Clienttxt.text
                VarLogSheet.Fields(4) = PaperWeighttxt.text
                VarLogSheet.Fields(5) = Val(qtytxt.text)
                
                VarLogSheet.Fields(6) = cmbPrintStatus.text
                VarLogSheet.Fields(7) = cmbPapertype.text
                VarLogSheet.Fields(8) = cmbMachine.text
                VarLogSheet.Fields(9) = cmbColorStatus.text
                VarLogSheet.Fields(10) = DailylogsheetDate
                
                VarLogSheet.Fields(11) = Pricetxt.text
                VarLogSheet.Fields(12) = qtytxt.text * Pricetxt.text
                
                
                VarLogSheet.Update
CommonDialog1.ShowPrinter
Printer.EndDoc

'--------------------------------------------------------------------------------------------'
'RECEIPT CODE'
'----------------------'
ShopNamelbl.Visible = False
ShopContactlbl.Visible = False
ShopAddresslbl.Visible = False
Cnamelbl.Visible = False
Pnamelbl.Visible = False
Qtylbl.Visible = False
Plbl.Visible = False
TClbl.Visible = False
'---------------------------------------------------------------------------------------------'
'RECEIPT LINES'
'----------------------'
line1lbl.Visible = False
Line2lbl.Visible = False
Line3lbl.Visible = False
Line4lbl.Visible = False
Line5lbl.Visible = False
Line6lbl.Visible = False
'---------------------------------------------------------------------------------------------'
'RECEIPT DETAILS'
'-------------------'
CustomerNamelbl.Visible = False
ProductNamelbl.Visible = False
Quantitylbl.Visible = False
Pricelbl.Visible = False
TotalCostlbl.Visible = False
DTlbl.Visible = False
Datelbl.Visible = False
Signature1lbl.Visible = False
Signature2.Visible = False
Signaturewhite.Visible = False

'--------------------------------------------------------------------------------------------'
Clear
cmdRecieptGenerate.Visible = True
cmdPrintReceipt.Visible = False
End If
End If

End Sub

Private Sub cmdRecieptGenerate_Click()
   If JobNametxt.text = "" Then
        If Sizetxt.text = "" Then
            If Clienttxt.text = "" Then
                If PaperWeighttxt.text = "" Then
                    If qtytxt.text = "" Then
                        If cmbPrintStatus.text = "Select Your Choice" Then
                            If cmbPapertype.text = "Select Your Choice" Then
                                If cmbMachine.text = "Select Your Choice" Then
                                    If cmbColorStatus.text = "Select Your Choice" Then
                                        If Pricetxt.text = "" Then
                                            MsgBox "Please Enter All The Details Properly", vbOKOnly + vbCritical, App.Title
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        If JobNametxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If Sizetxt.text = "" Then
           MsgBox "Please Enter Paper Size", vbOKOnly + vbCritical, App.Title
        End If
        
        If Clienttxt.text = "" Then
            MsgBox "Please Client Name", vbOKOnly + vbCritical, App.Title
        End If
        
        If PaperWeighttxt.text = "" Then
            MsgBox "Please Enter Paper Weight In GSM", vbOKOnly + vbCritical, App.Title
        End If
        
        If qtytxt.text = "" Then
            MsgBox "Please Enter Job Name", vbOKOnly + vbCritical, App.Title
        End If
    End If
    '--------------------------------------------------------------------------------------------'
'RECEIPT CODE'
'----------------------'
ShopNamelbl.Visible = True
ShopContactlbl.Visible = True
ShopAddresslbl.Visible = True
Cnamelbl.Visible = True
Pnamelbl.Visible = True
Qtylbl.Visible = True
Plbl.Visible = True
TClbl.Visible = True
'---------------------------------------------------------------------------------------------'
'RECEIPT LINES'
'----------------------'
line1lbl.Visible = True
Line2lbl.Visible = True
Line3lbl.Visible = True
Line4lbl.Visible = True
Line5lbl.Visible = True
Line6lbl.Visible = True
'---------------------------------------------------------------------------------------------'
'RECEIPT DETAILS'
'-------------------'
CustomerNamelbl.Visible = True
ProductNamelbl.Visible = True
Quantitylbl.Visible = True
Pricelbl.Visible = True
TotalCostlbl.Visible = True
DTlbl.Visible = True
Datelbl.Visible = True
Signature1lbl.Visible = True
Signature2.Visible = False
Signaturewhite.Visible = True
'---------------------------------------------------------------------------------------------'
'GENERATE RECEIPT BUTTON HIDE AND PRINT RECEIPT BUTTON SHOW'
'--------------------------------------------------------------'
cmdRecieptGenerate.Visible = False
cmdPrintReceipt.Visible = True

'---------------------------------------------------------------------------------------------'
'PRINTING DETAILS'
'-------------------'
CustomerNamelbl.Caption = Clienttxt.text
ProductNamelbl.Caption = cmbPapertype
Quantitylbl.Caption = qtytxt.text
Datelbl.Caption = DailylogsheetDate
Pricelbl.Caption = Pricetxt.text
TotalCostlbl.Caption = qtytxt.text * Pricetxt.text





End Sub







Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    
    DailyLogSheet.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
    
    cmbPrintStatus.AddItem ("Single Side")
    cmbPrintStatus.AddItem ("Double")
    
    cmbPapertype.AddItem ("Art Paper")
    cmbPapertype.AddItem ("Sticker")
    cmbPapertype.AddItem ("Mate")
    cmbPapertype.AddItem ("Glossy")
    
    cmbMachine.AddItem ("Konica Minolta")
    cmbMachine.AddItem ("Canon Image Runner")
    
    cmbColorStatus.AddItem ("Black And White")
    cmbColorStatus.AddItem ("Coloring")
    
   
    
    
 
'--------------------------------------------------------------------------------------------'
'RECEIPT CODE'
'----------------------'
ShopNamelbl.Visible = False
ShopContactlbl.Visible = False
ShopAddresslbl.Visible = False
Cnamelbl.Visible = False
Pnamelbl.Visible = False
Qtylbl.Visible = False
Plbl.Visible = False
TClbl.Visible = False
'---------------------------------------------------------------------------------------------'
'RECEIPT LINES'
'----------------------'
line1lbl.Visible = False
Line2lbl.Visible = False
Line3lbl.Visible = False
Line4lbl.Visible = False
Line5lbl.Visible = False
Line6lbl.Visible = False
'---------------------------------------------------------------------------------------------'
'RECEIPT DETAILS'
'-------------------'
CustomerNamelbl.Visible = False
ProductNamelbl.Visible = False
Quantitylbl.Visible = False
Pricelbl.Visible = False
TotalCostlbl.Visible = False
DTlbl.Visible = False
Datelbl.Visible = False
Signature1lbl.Visible = False
Signature2.Visible = False
Signaturewhite.Visible = False

'--------------------------------------------------------------------------------------------'
'PRINT BUTTON HIDE'
'---------------------'
cmdPrintReceipt.Visible = False



'--------------------------------------------------------------------------------------------'
    

End Sub
Private Sub Form_Resize()
    DailyLogSheet.Width = Me.ScaleWidth
    DailyLogSheet.Height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub




Private Sub JobNametxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sizetxt.SetFocus
    End If
End Sub

Private Sub PaperWeighttxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        qtytxt.SetFocus
    End If
End Sub

Private Sub qtytxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPrintStatus.SetFocus
    End If
End Sub

Private Sub Sizetxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Clienttxt.SetFocus
    End If
End Sub


Private Sub Clienttxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PaperWeighttxt.SetFocus
    End If
End Sub

Private Sub cmbColorStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DailylogsheetDate.SetFocus
    End If
End Sub

Private Sub cmbMachine_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbColorStatus.SetFocus
    End If
End Sub

Private Sub cmbPapertype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbMachine.SetFocus
    End If
End Sub

Private Sub cmbPrintStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPapertype.SetFocus
    End If
End Sub
