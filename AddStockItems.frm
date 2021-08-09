VERSION 5.00
Begin VB.Form PurchasedItems 
   Caption         =   "Add Items"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   17835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame AddFrame 
      BackColor       =   &H00808080&
      Caption         =   "ADD ITEMS"
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
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   17775
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NO."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchased Items"
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
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   9615
   End
   Begin VB.Image AddStocksImage 
      Height          =   8055
      Left            =   0
      Picture         =   "AddStockItems.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17775
   End
End
Attribute VB_Name = "PurchasedItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
