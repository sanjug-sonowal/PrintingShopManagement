VERSION 5.00
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Printing Shop Management System"
   ClientHeight    =   8070
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15960
   LinkTopic       =   "M"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu Home 
      Caption         =   "&Home"
   End
   Begin VB.Menu Inventory 
      Caption         =   "&Inventory"
      Begin VB.Menu StockDetails 
         Caption         =   "Stock Details"
         Shortcut        =   ^S
      End
      Begin VB.Menu Purchase 
         Caption         =   "Purchase"
         Shortcut        =   ^P
      End
      Begin VB.Menu Supplier_Details 
         Caption         =   "Supplier Details"
         Shortcut        =   ^U
      End
      Begin VB.Menu Sales 
         Caption         =   "Sales"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
   Begin VB.Menu Exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit








Private Sub About_Click()
    Unload DailyLogSheet
    Unload Supplier_Info
    Unload SalesForm
    Unload Stocks
    Unload PurchasedItem
    AboutForm1.Show
End Sub

Private Sub Help_Click()
 Unload DailyLogSheet
    Unload Supplier_Info
    Unload SalesForm
    Unload Stocks
    Unload PurchasedItem
    Unload AboutForm1
    HelpForm.Show
    
    

End Sub

Private Sub Home_Click()
    
    Unload Supplier_Info
    Unload SalesForm
    Unload Stocks
    Unload PurchasedItem
    Unload AboutForm1
    
    DailyLogSheet.Show
    
End Sub

Private Sub MDIForm_Load()
    Call Home_Click
    
    PurchasedItem.Move 0, 0, Me.Width, Me.Height
    DailyLogSheet.Move 0, 0, Me.Width, Me.Height
    Stocks.Move 0, 0, Me.Width, Me.Height
    Supplier_Info.Move 0, 0, Me.Width, Me.Height
    SalesForm.Move 0, 0, Me.Width, Me.Height
    AboutForm1.Move 0, 0, Me.Width, Me.Height
   
    
    
    
    
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
    Me.WindowState = 2
    
End Sub


Private Sub MDIForm_Resize()
    DailyLogSheet.Width = Me.ScaleWidth
    DailyLogSheet.Height = Me.ScaleHeight
    
    Stocks.Width = Me.ScaleWidth
    Stocks.Height = Me.ScaleHeight
    
    PurchasedItem.Width = Me.ScaleWidth
    PurchasedItem.Height = Me.ScaleHeight
    
    Supplier_Info.Width = Me.ScaleWidth
    Supplier_Info.Height = Me.ScaleHeight
    
    SalesForm.Width = Me.ScaleWidth
    SalesForm.Height = Me.ScaleHeight
    
    AboutForm1.Width = Me.ScaleWidth
    AboutForm1.Height = Me.ScaleHeight
    
   
    
    
   
    
    
    
    ResizeControls Me, initialcontrollist
End Sub


Private Sub Purchase_Click()
DailyLogSheet.Visible = False
Stocks.Visible = False


PurchasedItem.Show


End Sub

Private Sub Sales_Click()
PurchasedItem.Visible = False
Stocks.Visible = False


DailyLogSheet.Visible = False

PurchasedItem.Visible = False

Supplier_Info.Visible = False

SalesForm.Show


End Sub

Private Sub StockDetails_Click()
DailyLogSheet.Visible = False

PurchasedItem.Visible = False

Supplier_Info.Visible = False

SalesForm.Visible = False

AboutForm1.Visible = False

Stocks.Show




End Sub

Private Sub Supplier_Details_Click()
Supplier_Info.Show
DailyLogSheet.Visible = False
Stocks.Visible = False
PurchasedItem.Visible = False
End Sub
