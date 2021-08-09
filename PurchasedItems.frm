VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchasedItems 
   ClientHeight    =   8160
   ClientLeft      =   -60
   ClientTop       =   120
   ClientWidth     =   17835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   17835
   Begin VB.CommandButton cmdHome_btn 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   24
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmcdBack_btn 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   23
      Top             =   7440
      Width           =   2415
   End
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
      Begin VB.TextBox Qty_txt 
         Height          =   495
         Left            =   11040
         TabIndex        =   36
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox PaperWeight_txt 
         Height          =   495
         Left            =   6720
         TabIndex        =   34
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox Description_txt 
         Height          =   495
         Left            =   6720
         TabIndex        =   30
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox Brand_txt 
         Height          =   495
         Left            =   6720
         TabIndex        =   29
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox SupplierId_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   27
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox ProductName_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   25
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Caption         =   "CONTROL PANEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   6000
         TabIndex        =   16
         Top             =   5040
         Width           =   11775
         Begin VB.CommandButton cmdExit 
            Caption         =   "EXIT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9720
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7800
            TabIndex        =   21
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "UPDATE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5880
            TabIndex        =   20
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "NEXT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   19
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdPrev 
            Caption         =   "PREVIOUS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   18
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "ADD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1935
         End
      End
      Begin MSComCtl2.DTPicker DtPickerPurchasedItems 
         Height          =   495
         Left            =   6720
         TabIndex        =   15
         Top             =   4320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   118882305
         CurrentDate     =   44153
      End
      Begin VB.TextBox Category_txt 
         Height          =   495
         Left            =   6720
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox SupplierAddress_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox SupplierMob_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   4920
         Width           =   2415
      End
      Begin VB.TextBox SupplierName_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox ProdductId_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Invoice_txt 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
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
         Height          =   615
         Left            =   9840
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PAPER WEIGHT (GSM)"
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
         Height          =   855
         Left            =   5400
         TabIndex        =   33
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER ADDRESS"
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
         TabIndex        =   32
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER MOB NO."
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
         TabIndex        =   31
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER NAME"
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
         TabIndex        =   28
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER ID"
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
         TabIndex        =   26
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   3375
         Left            =   10800
         Picture         =   "PurchasedItems.frx":0000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
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
         Left            =   5400
         TabIndex        =   14
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT NAME"
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
         TabIndex        =   13
         Top             =   2520
         Width           =   1455
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
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION"
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
         Left            =   5400
         TabIndex        =   9
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "BRAND"
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
         Left            =   5400
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY"
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
         Left            =   5400
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT ID"
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
         TabIndex        =   3
         Top             =   1680
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
      TabIndex        =   2
      Top             =   120
      Width           =   9615
   End
   Begin VB.Image AddStocksImage 
      Height          =   8055
      Left            =   0
      Picture         =   "PurchasedItems.frx":3FE6
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
Private initialcontrollist() As ControlInitial
Option Explicit



Private Sub cmcdBack_btn_Click()
    Stocks.Show
End Sub

Private Sub cmdAdd_Click()
  If Invoice_txt.Text = "" Then
    If ProdductId_txt.Text = "" Then
        If ProductName_txt.Text = "" Then
            If SupplierId_txt.Text = "" Then
                If SupplierName_txt.Text = "" Then
                    If SupplierMob_txt.Text = "" Then
                        If SupplierAddress_txt.Text = "" Then
                            If Category_txt.Text = "" Then
                                If Brand_txt.Text = "" Then
                                    If Description_txt.Text = "" Then
                                        If PaperWeight_txt.Text = "" Then
                                            If Qty_txt.Text = "" Then
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
        End If
    End If
    
    If Invoice_txt.Text = "" Then
            MsgBox "Please Enter Invoice Number", vbOKOnly + vbCritical, App.Title
    End If
    
    If ProdductId_txt.Text = "" Then
            MsgBox "Please Enter Product ID", vbOKOnly + vbCritical, App.Title
    End If
    
    If ProductName_txt.Text = "" Then
            MsgBox "Please Enter Product Name", vbOKOnly + vbCritical, App.Title
    End If
    
    If SupplierId_txt.Text = "" Then
            MsgBox "Please Enter Supplier ID", vbOKOnly + vbCritical, App.Title
    End If
    
    If SupplierName_txt.Text = "" Then
            MsgBox "Please Enter Supplier Name", vbOKOnly + vbCritical, App.Title
    End If
    
    If SupplierMob_txt.Text = "" Then
            MsgBox "Please Enter Supplier Mobile Number", vbOKOnly + vbCritical, App.Title
    End If
    
    If SupplierAddress_txt.Text = "" Then
            MsgBox "Please Enter Supplier Address", vbOKOnly + vbCritical, App.Title
    End If
    
    If Category_txt.Text = "" Then
            MsgBox "Please Enter Category", vbOKOnly + vbCritical, App.Title
    End If
    
    If Brand_txt.Text = "" Then
            MsgBox "Please Enter Brand", vbOKOnly + vbCritical, App.Title
    End If
    
    If Description_txt.Text = "" Then
            MsgBox "Please Enter Description", vbOKOnly + vbCritical, App.Title
    End If
    
    If PaperWeight_txt.Text = "" Then
            MsgBox "Please Enter Paper Weight in (GSM)", vbOKOnly + vbCritical, App.Title
    End If
    
    If Qty_txt.Text = "" Then
            MsgBox "Please Enter Quantity", vbOKOnly + vbCritical, App.Title
    End If
    
    Else
    
    
    
   
    
    '----------------------------------------------------------------------------------------------------'
    'ADD TO STOCK DETAILS'
    '----------------------------------------------------------------------------------------------------'
    
    VarStockDetails.AddNew
    VarStockDetails.Fields(0) = Invoice_txt.Text
    VarStockDetails.Fields(1) = ProdductId_txt.Text
    VarStockDetails.Fields(2) = ProductName_txt.Text
    VarStockDetails.Fields(3) = PaperWeight_txt.Text
    VarStockDetails.Fields(4) = Brand_txt.Text
    VarStockDetails.Fields(5) = Category_txt.Text
    VarStockDetails.Fields(6) = Description_txt.Text
    VarStockDetails.Fields(7) = DtPickerPurchasedItems
    VarStockDetails.Fields(8) = Qty_txt.Text
    VarStockDetails.Update
    
    
    '-----------------------------------------------------------------------------------------------------'
    '-----------------------------------------------------------------------------------------------------'
    'ADD TO SUPPLIER DETAILS'
    '-----------------------------------------------------------------------------------------------------'
    
    VarSupplierDetails.AddNew
    VarSupplierDetails.Fields(0) = SupplierId_txt.Text
    VarSupplierDetails.Fields(1) = SupplierName_txt.Text
    VarSupplierDetails.Fields(2) = SupplierMob_txt.Text
    VarSupplierDetails.Fields(3) = SupplierAddress_txt.Text
    VarSupplierDetails.Update
    
            
    '-----------------------------------------------------------------------------------------------------'
    '-----------------------------------------------------------------------------------------------------'
    'ADD TO PURCHASED DETAILS'
    '-----------------------------------------------------------------------------------------------------'
     
     VarPurchasedDetails.AddNew
     VarPurchasedDetails.Fields(0) = VarStockDetails(0)
     VarPurchasedDetails.Fields(1) = VarStockDetails(1)
     VarPurchasedDetails.Fields(2) = VarStockDetails(2)
     VarPurchasedDetails.Fields(3) = VarStockDetails(3)
     
     VarPurchasedDetails.Fields(4) = VarSupplierDetails(0)
     VarPurchasedDetails.Fields(5) = VarSupplierDetails(1)
     VarPurchasedDetails.Fields(6) = VarSupplierDetails(2)
     VarPurchasedDetails.Fields(7) = VarSupplierDetails(3)
     
     VarPurchasedDetails.Fields(8) = VarStockDetails(5)
     VarPurchasedDetails.Fields(9) = VarStockDetails(4)
     VarPurchasedDetails.Fields(10) = VarStockDetails(6)
     VarPurchasedDetails.Fields(11) = VarStockDetails(8)
     VarPurchasedDetails.Fields(12) = VarStockDetails(7)
     VarPurchasedDetails.Update
     
     
     MsgBox "record saved successfully"
     '-----------------------------------------------------------------------------------------------------'
    
End If
    
    
                                
            
End Sub

Private Sub cmdHome_btn_Click()
    DailyLogSheet.Show
End Sub

Private Sub Form_Load()
    DailyLogSheet.Visible = False
    PurchasedItems.Move 0, 0, Me.width, Me.height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.height, Me.width, Me.Left, Me.Top
End Sub

Private Sub Form_Resize()
    PurchasedItems.width = Me.ScaleWidth
    PurchasedItems.height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub
