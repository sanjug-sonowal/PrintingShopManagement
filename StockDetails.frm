VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Stocks 
   BorderStyle     =   0  'None
   Caption         =   "Stock Details"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   12960
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmbNewEntry 
      Caption         =   "New Entry"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Frame StockDetailsFrame 
      BackColor       =   &H00808080&
      Caption         =   "STOCK DETAILS"
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
      Height          =   6255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   16215
      Begin MSDataGridLib.DataGrid StockDetailsGrid 
         Height          =   5415
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   9551
         _Version        =   393216
         BackColor       =   4210752
         ForeColor       =   16777215
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK DETAILS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   1815
      Left            =   5520
      TabIndex        =   5
      Top             =   -360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3201
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Image StockDetailsImage 
      Height          =   7935
      Left            =   0
      Picture         =   "StockDetails.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16215
   End
End
Attribute VB_Name = "Stocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit





Private Sub cmbNewEntry_Click()
PurchasedItem.Show
End Sub

Private Sub cmdGenerateReport_Click()
StockReport.Show
End Sub

Private Sub Form_Load()
    Dim VarStock As New Recordset
    VarStock.Open "select * from StockDetails", cn, adOpenDynamic, adLockOptimistic
    Set StockDetailsGrid.DataSource = VarStock
    StockDetailsGrid.Refresh
    Set VarStock = Nothing
    Stocks.Move 0, 0, Me.Width, Me.Height
    
    
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
    
    
    
    cmbNewEntry.FontBold = True
    cmbNewEntry.FontName = "Tahoma"
    
   
    
    cmdGenerateReport.FontBold = True
    cmdGenerateReport.FontName = "Tahoma"
    
  
    
    
End Sub
Private Sub Form_Resize()
    Stocks.Width = Me.ScaleWidth
    Stocks.Height = Me.ScaleHeight
    
    
    ResizeControls Me, initialcontrollist
End Sub
