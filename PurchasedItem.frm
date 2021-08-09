VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PurchasedItem 
   BorderStyle     =   0  'None
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20415
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   20415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Print Report"
      Height          =   495
      Left            =   240
      TabIndex        =   39
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton Home_btn 
      Caption         =   "HOME"
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
      Left            =   18480
      TabIndex        =   7
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "ADD ITEMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   20415
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Left            =   8040
         TabIndex        =   42
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1560
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   6960
         Width           =   2655
      End
      Begin VB.TextBox Search_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11640
         TabIndex        =   38
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo_Search 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "PurchasedItem.frx":0000
         Left            =   9600
         List            =   "PurchasedItem.frx":0010
         TabIndex        =   37
         Text            =   "Select Your Choice"
         Top             =   360
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid PurchasedGrid 
         Height          =   5175
         Left            =   9600
         TabIndex        =   36
         Top             =   840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   9128
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
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
      Begin VB.TextBox Price_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   35
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox Quantity_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   33
         Top             =   3960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   6840
         TabIndex        =   31
         Top             =   5640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         Format          =   118423553
         CurrentDate     =   44204
      End
      Begin VB.TextBox Paper_Weight_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   30
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Description_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   29
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Brand_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   28
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Category_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   27
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Supplier_Address_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   26
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox Supplier_Mob_No_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   25
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox Supplier_Name_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   24
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox Supplier_Id_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   23
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Product_Name_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   22
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Product_Id_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Invoice_No_txt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "CONTROL PANEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   10200
         TabIndex        =   2
         Top             =   6240
         Width           =   9495
         Begin VB.CommandButton Exit_btn 
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
            Height          =   615
            Left            =   7080
            TabIndex        =   6
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Delete_btn 
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
            Height          =   615
            Left            =   4800
            TabIndex        =   5
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Update_btn 
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
            Height          =   615
            Left            =   2520
            TabIndex        =   4
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Add_btn 
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
            Height          =   615
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE"
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
         Height          =   495
         Left            =   5040
         TabIndex        =   34
         Top             =   4920
         Width           =   1455
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
         Height          =   495
         Left            =   5040
         TabIndex        =   32
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PAPER WEIGHT"
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
         Height          =   495
         Left            =   5040
         TabIndex        =   19
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label12 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   18
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   17
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label10 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label9 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
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
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label7 
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
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
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
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
         Height          =   7095
         Left            =   -480
         TabIndex        =   40
         Top             =   -120
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   12515
         BackColor       =   4210752
         MousePointer    =   0
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchased Item"
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
      Height          =   975
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   8535
   End
   Begin VB.Image PurchasedImage 
      Height          =   9495
      Left            =   0
      Picture         =   "PurchasedItem.frx":0045
      Stretch         =   -1  'True
      Top             =   120
      Width           =   20415
   End
End
Attribute VB_Name = "PurchasedItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit
Private Function Clear()
Invoice_No_txt.text = ""
Product_Id_txt.text = ""
Product_Name_txt.text = ""
Supplier_Id_txt.text = ""
Supplier_Name_txt.text = ""
Supplier_Mob_No_txt.text = ""
Supplier_Address_txt.text = ""
Category_txt.text = ""
Brand_txt.text = ""
Description_txt.text = ""
Paper_Weight_txt.text = ""
Quantity_txt.text = ""
Price_txt.text = ""
End Function
Private Sub Add_btn_Click()
If Product_Id_txt.text = "" Then
        If Product_Name_txt.text = "" Then
            If Supplier_Id_txt.text = "" Then
                If Supplier_Name_txt.text = "" Then
                    If Supplier_Mob_No_txt.text = "" Then
                        If Supplier_Address_txt.text = "" Then
                            If Category_txt.text = "" Then
                                If Brand_txt.text = "" Then
                                    If Description_txt.text = "" Then
                                        If Paper_Weight_txt.text = "" Then
                                            If Quantity_txt.text = "" Then
                                                If Price_txt.text = "" Then
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
    
    
  
    
    If Product_Id_txt.text = "" Then
            MsgBox "Please Enter Product ID", vbOKOnly + vbCritical, App.Title
    End If
    
    If Product_Name_txt.text = "" Then
            MsgBox "Please Enter Product Name", vbOKOnly + vbCritical, App.Title
    End If
    
    If Supplier_Id_txt.text = "" Then
            MsgBox "Please Enter Supplier ID", vbOKOnly + vbCritical, App.Title
    End If
    
    If Supplier_Name_txt.text = "" Then
            MsgBox "Please Enter Supplier Name", vbOKOnly + vbCritical, App.Title
    End If
    
    If Supplier_Mob_No_txt.text = "" Then
            MsgBox "Please Enter Supplier Mobile Number", vbOKOnly + vbCritical, App.Title
    End If
    
    If Supplier_Address_txt.text = "" Then
            MsgBox "Please Enter Supplier Address", vbOKOnly + vbCritical, App.Title
    End If
    
    If Category_txt.text = "" Then
            MsgBox "Please Enter Category", vbOKOnly + vbCritical, App.Title
    End If
    
    If Brand_txt.text = "" Then
            MsgBox "Please Enter Brand", vbOKOnly + vbCritical, App.Title
    End If
    
    If Description_txt.text = "" Then
            MsgBox "Please Enter Description", vbOKOnly + vbCritical, App.Title
    End If
    
    If Paper_Weight_txt.text = "" Then
            MsgBox "Please Enter Paper Weight in (GSM)", vbOKOnly + vbCritical, App.Title
    End If
    
    If Quantity_txt.text = "" Then
            MsgBox "Please Enter Quantity", vbOKOnly + vbCritical, App.Title
    End If
    
    If Price_txt.text = "" Then
            MsgBox "Please Enter Price", vbOKOnly + vbCritical, App.Title
    End If
    
    Else
    
    
    
   
    
    '----------------------------------------------------------------------------------------------------'
    'ADD TO STOCK DETAILS'
    '----------------------------------------------------------------------------------------------------'
 
    
    
    '-----------------------------------------------------------------------------------------------------'
    '-----------------------------------------------------------------------------------------------------'
    'ADD TO SUPPLIER DETAILS'
    '-----------------------------------------------------------------------------------------------------'
 
            
    '-----------------------------------------------------------------------------------------------------'
    '-----------------------------------------------------------------------------------------------------'
    'ADD TO PURCHASED DETAILS'
    '-----------------------------------------------------------------------------------------------------'
    VarPurchasedDetails.AddNew

            VarPurchasedDetails.Fields(0) = Val(Invoice_No_txt.text)
     VarPurchasedDetails.Fields(1) = Val(Product_Id_txt.text)
     VarPurchasedDetails.Fields(2) = Product_Name_txt.text
     VarPurchasedDetails.Fields(3) = Val(Paper_Weight_txt.text)
     
     VarPurchasedDetails.Fields(4) = Val(Supplier_Id_txt.text)
     VarPurchasedDetails.Fields(5) = Supplier_Name_txt.text
     VarPurchasedDetails.Fields(6) = Val(Supplier_Mob_No_txt.text)
     VarPurchasedDetails.Fields(7) = Supplier_Address_txt.text
     
     VarPurchasedDetails.Fields(8) = Category_txt.text
     VarPurchasedDetails.Fields(9) = Brand_txt.text
     VarPurchasedDetails.Fields(10) = Description_txt.text
     VarPurchasedDetails.Fields(11) = Val(Quantity_txt.text)
     VarPurchasedDetails.Fields(12) = Val(Price_txt.text)
     VarPurchasedDetails.Fields(13) = VarPurchasedDetails.Fields(11) * VarPurchasedDetails.Fields(12)
     VarPurchasedDetails.Fields(14) = DTPicker1.Value
            
            VarPurchasedDetails.Update
            VarPurchasedDetails.MoveFirst
            
            VarSupplierDetails.AddNew
            VarSupplierDetails.Fields(0) = Val(Invoice_No_txt.text)
    VarSupplierDetails.Fields(1) = Supplier_Id_txt.text
    VarSupplierDetails.Fields(2) = Supplier_Name_txt.text
    VarSupplierDetails.Fields(3) = Supplier_Mob_No_txt.text
    VarSupplierDetails.Fields(4) = Supplier_Address_txt.text
            
            VarSupplierDetails.Update
            
            VarSupplierDetails.MoveFirst
            VarStockDetails.AddNew
            VarStockDetails.Fields(0) = Val(Invoice_No_txt.text)
            VarStockDetails.Fields(1) = Val(Product_Id_txt.text)
            VarStockDetails.Fields(2) = Product_Name_txt.text
            VarStockDetails.Fields(3) = Val(Paper_Weight_txt.text)
            VarStockDetails.Fields(4) = Brand_txt.text
            VarStockDetails.Fields(5) = Category_txt.text
            VarStockDetails.Fields(6) = Description_txt.text
            VarStockDetails.Fields(7) = Val(Price_txt.text)
            VarStockDetails.Fields(8) = DTPicker1.Value
            VarStockDetails.Fields(9) = Val(Quantity_txt.text)
            
    
            
            VarStockDetails.Update
            
         


     MsgBox "record saved successfully"
     DisplayDBRecords
     Clear
     '-----------------------------------------------------------------------------------------------------'
    
End If
    
    
                                
            
End Sub



Private Sub Command1_Click()
If Combo_Search.text = "Invoice No" Then
VarProductIdUpdate.Open "select * from PurchasedDetails where Invoice_No like '%" & Search_txt.text & "%'", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarPurchasedDetails
PurchasedGrid.Refresh
Set VarPurchasedDetails = Nothing

End If
VarProductIdUpdate.Close
End Sub

Private Sub Command2_Click()
VarPurchasedDetails.AddNew

            VarPurchasedDetails.Fields(0) = Val(Invoice_No_txt.text)
     VarPurchasedDetails.Fields(1) = Val(Product_Id_txt.text)
     VarPurchasedDetails.Fields(2) = Product_Name_txt.text
     VarPurchasedDetails.Fields(3) = Val(Paper_Weight_txt.text)
     
     VarPurchasedDetails.Fields(4) = Val(Supplier_Id_txt.text)
     VarPurchasedDetails.Fields(5) = Supplier_Name_txt.text
     VarPurchasedDetails.Fields(6) = Val(Supplier_Mob_No_txt.text)
     VarPurchasedDetails.Fields(7) = Supplier_Address_txt.text
     
     VarPurchasedDetails.Fields(8) = Category_txt.text
     VarPurchasedDetails.Fields(9) = Brand_txt.text
     VarPurchasedDetails.Fields(10) = Description_txt.text
     VarPurchasedDetails.Fields(11) = Val(Quantity_txt.text)
     VarPurchasedDetails.Fields(12) = Val(Price_txt.text)
     VarPurchasedDetails.Fields(13) = VarPurchasedDetails.Fields(11) * VarPurchasedDetails.Fields(12)
     VarPurchasedDetails.Fields(14) = DTPicker1.Value
            
            VarPurchasedDetails.Update
            VarPurchasedDetails.MoveFirst
            
            VarSupplierDetails.AddNew
            VarSupplierDetails.Fields(0) = Val(Invoice_No_txt.text)
    VarSupplierDetails.Fields(1) = Supplier_Id_txt.text
    VarSupplierDetails.Fields(2) = Supplier_Name_txt.text
    VarSupplierDetails.Fields(3) = Supplier_Mob_No_txt.text
    VarSupplierDetails.Fields(4) = Supplier_Address_txt.text
            
            VarSupplierDetails.Update
            
            VarSupplierDetails.MoveFirst
            VarStockDetails.AddNew
            VarStockDetails.Fields(0) = Val(Invoice_No_txt.text)
            VarStockDetails.Fields(1) = Val(Product_Id_txt.text)
            VarStockDetails.Fields(2) = Product_Name_txt.text
            VarStockDetails.Fields(3) = Val(Paper_Weight_txt.text)
            VarStockDetails.Fields(4) = Brand_txt.text
            VarStockDetails.Fields(5) = Category_txt.text
            VarStockDetails.Fields(6) = Description_txt.text
            VarStockDetails.Fields(7) = Val(Price_txt.text)
            VarStockDetails.Fields(8) = DTPicker1.Value
            VarStockDetails.Fields(9) = Val(Quantity_txt.text)
            
    
            
            VarStockDetails.Update
            
         


MsgBox "success"


End Sub

Private Sub Command3_Click()
VarPurchasedDetails.Fields(0) = Val(Invoice_No_txt.text)
End Sub

Private Sub Delete_btn_Click()
Dim invoice As Integer
invoice = InputBox("Enter Invoice Number Which You Want to Delete From Database")

'---------------------------------------------------------------------------------------------------------------------
'DELETE FROM PURCHASED DETAILS TABLE IN THE DATABASE
'---------------------------------------------------------------------------------------------------------------------

VarRecordDelete.Open "select * from PurchasedDetails where Invoice_No=" & invoice, cn, adOpenDynamic, adLockOptimistic
If Not (VarRecordDelete.EOF) Then
VarRecordDelete.Delete
End If

VarRecordDelete.Close
DisplayDBRecords
'--------------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------------------------
'DELETE FROM SUPPLIER DETAILS TABLE IN THE DATABASE
'--------------------------------------------------------------------------------------------------------------------

VarRecordDelete.Open "select * from SupplierDetails where Invoice_No=" & invoice, cn, adOpenDynamic, adLockOptimistic
If Not (VarRecordDelete.EOF) Then
VarRecordDelete.Delete
End If
VarRecordDelete.Close
DisplayDBRecords
'--------------------------------------------------------------------------------------------------------------------
'DELETE FROM STOCK DETAILS TABLE INTHE DATABASE
'--------------------------------------------------------------------------------------------------------------------

VarRecordDelete.Open "select * from StockDetails where Invoice_No=" & invoice, cn, adOpenDynamic, adLockOptimistic
If Not (VarRecordDelete.EOF) Then
VarRecordDelete.Delete
MsgBox "Record Was Deleted Successfully"
Else
MsgBox "Record Not Found In The Database"
End If

DisplayDBRecords

End Sub

Private Sub Exit_btn_Click()
Unload Me
DailyLogSheet.Show
End Sub
Private Function DisplayDBRecords()
Dim VarPur As New Recordset
VarPur.Open "select * from PurchasedDetails", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarPur
PurchasedGrid.Refresh
Set VarPur = Nothing
End Function
Private Sub Form_Load()


DisplayDBRecords

'------------------------------------------------------------------





'------------------------------------------------------------------

PurchasedImage.Move 0, 0, Me.Width, Me.Height

initialcontrollist = GetLocation(Me)
ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub

Private Sub Form_Resize()
    PurchasedImage.Width = Me.ScaleWidth
    PurchasedImage.Height = Me.ScaleHeight
    
    ResizeControls Me, initialcontrollist
End Sub
Private Function Display()
Invoice_No_txt = VarPurchasedDetails(0)
Product_Id_txt = VarPurchasedDetails(1)
Product_Name_txt = VarPurchasedDetails(2)
Supplier_Id_txt = VarPurchasedDetails(3)
Supplier_Name_txt = VarPurchasedDetails(4)
Supplier_Mob_No_txt = VarPurchasedDetails(5)
Supplier_Address_txt = VarPurchasedDetails(6)
Category_txt = VarPurchasedDetails(7)
Brand_txt = VarPurchasedDetails(8)
Description_txt = VarPurchasedDetails(9)
Paper_Weight_txt = VarPurchasedDetails(10)
Quantity_txt = VarPurchasedDetails(11)
Price_txt = VarPurchasedDetails(12)
DTPicker1 = VarPurchasedDetails(14)
End Function

Private Sub Next_btn_Click()

  VarPurchasedDetails.MoveNext
If Not VarPurchasedDetails.EOF Then
Display
Else
VarPurchasedDetails.MoveFirst
Display
End If
End Sub

Private Sub Previous_btn_Click()

VarPurchasedDetails.MovePrevious
If VarPurchasedDetails.BOF Then
VarPurchasedDetails.MoveLast
Display
Else
Display
End If

End Sub







Private Sub PurchasedGrid_Click()

'-----------------------------------------------------------------------------------------
'Variable's For Displaying Selected Records from DataGrid
'-----------------------------------------------------------------------------------------
 
 Dim Invoice_No As Integer
 Dim Product_Id As Integer
 Dim Product_Name As String
 Dim Paper_Weight As Integer
 Dim Supplier_Id As Integer
 Dim supplier_name As String
 Dim Supplier_Mob_No As String
 Dim Supplier_Address As String
 Dim category As String
 Dim brand As String
 Dim Description As String
 Dim Quantity As Integer
 Dim Price As Integer
 Dim P_Date As String
 
 '----------------------------------------------------------------------------------------
 'Fetching Data Into The Textbox From the Database
 '----------------------------------------------------------------------------------------
 
 Invoice_No = PurchasedGrid.Columns(0)
 Product_Id = PurchasedGrid.Columns(1)
 Product_Name = PurchasedGrid.Columns(2)
 Paper_Weight = PurchasedGrid.Columns(3)
 Supplier_Id = PurchasedGrid.Columns(4)
 supplier_name = PurchasedGrid.Columns(5)
 Supplier_Mob_No = PurchasedGrid.Columns(6)
 Supplier_Address = PurchasedGrid.Columns(7)
 category = PurchasedGrid.Columns(8)
 brand = PurchasedGrid.Columns(9)
 Description = PurchasedGrid.Columns(10)
 Quantity = PurchasedGrid.Columns(11)
 Price = PurchasedGrid.Columns(12)
 P_Date = PurchasedGrid.Columns(14)
 
 '----------------------------------------------------------------------------------------
 'Displaying Selected Records from DataGrid Into the Textbox
 '----------------------------------------------------------------------------------------
 
 Invoice_No_txt.text = Invoice_No
 Product_Id_txt.text = Product_Id
 Product_Name_txt.text = Product_Name
 Supplier_Id_txt.text = Supplier_Id
 Supplier_Name_txt.text = supplier_name
 Supplier_Mob_No_txt.text = Supplier_Mob_No
 Supplier_Address_txt.text = Supplier_Address
 Category_txt.text = category
 Brand_txt.text = brand
 Description_txt.text = Description
 Paper_Weight_txt.text = Paper_Weight
 Quantity_txt.text = Quantity
 Price_txt.text = Price
 DTPicker1.Value = P_Date

'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------

End Sub



Private Sub search_Click()

End Sub
Private Function category()
Set VarSupplierIdUpdate = New ADODB.Recordset

VarSupplierIdUpdate.CursorLocation = adUseClient
If Combo_Search.ListIndex = 3 Then
 VarSupplierIdUpdate.Open "select * from PurchasedDetails where Category like '%" & Search_txt.text & "%'", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarSupplierIdUpdate
VarSupplierIdUpdate.ActiveConnection = Nothing
End If
End Function

Private Function brand()
Set VarSupplierIdUpdate = New ADODB.Recordset

VarSupplierIdUpdate.CursorLocation = adUseClient
If Combo_Search.ListIndex = 1 Then
 VarSupplierIdUpdate.Open "select * from PurchasedDetails where Brand like '%" & Search_txt.text & "%'", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarSupplierIdUpdate
VarSupplierIdUpdate.ActiveConnection = Nothing
End If
End Function
Private Function invoice()
Set VarSupplierIdUpdate = New ADODB.Recordset

VarSupplierIdUpdate.CursorLocation = adUseClient
If Combo_Search.ListIndex = 0 Then
 VarSupplierIdUpdate.Open "select * from PurchasedDetails where Invoice_No like '%" & Search_txt.text & "%'", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarSupplierIdUpdate
VarSupplierIdUpdate.ActiveConnection = Nothing
End If
End Function
Private Function supplier_name()
If Combo_Search.ListIndex = 2 Then
 VarSupplierIdUpdate.Open "select * from PurchasedDetails where Supplier_Name like '%" & Search_txt.text & "%'", cn, adOpenDynamic, adLockOptimistic
Set PurchasedGrid.DataSource = VarSupplierIdUpdate
VarSupplierIdUpdate.ActiveConnection = Nothing
End If
End Function
Private Sub Search_txt_Change()
invoice
brand
supplier_name
category

End Sub

Private Sub Update_btn_Click()

VarInvoiceUpdate.Open "Select * from PurchasedDetails where Invoice_No=" & Val(Invoice_No_txt.text) & "", cn, adOpenDynamic, adLockOptimistic
If VarInvoiceUpdate.EOF Then
MsgBox "Data Not Found"
VarInvoiceUpdate.Close
Else
VarInvoiceUpdate(1) = Product_Id_txt.text
VarInvoiceUpdate(2) = Product_Name_txt.text
VarInvoiceUpdate(3) = Paper_Weight_txt.text
VarInvoiceUpdate(4) = Supplier_Id_txt.text
VarInvoiceUpdate(5) = Supplier_Name_txt.text
VarInvoiceUpdate(6) = Supplier_Mob_No_txt.text
VarInvoiceUpdate(7) = Supplier_Address_txt.text
VarInvoiceUpdate(8) = Category_txt.text
VarInvoiceUpdate(9) = Brand_txt.text
VarInvoiceUpdate(10) = Description_txt.text
VarInvoiceUpdate(11) = Quantity_txt.text
VarInvoiceUpdate(12) = Price_txt.text
VarInvoiceUpdate(13) = DTPicker1
VarInvoiceUpdate.Update



MsgBox "Updated Successfully"
DisplayDBRecords

Clear

VarInvoiceUpdate.Close
End If




End Sub
