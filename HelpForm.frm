VERSION 5.00
Begin VB.Form HelpForm 
   Caption         =   "Help Form"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   21735
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   10455
      Left            =   120
      Picture         =   "HelpForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21495
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit

Private Sub Form_Load()
    HelpForm.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
    Unload DailyLogSheet
    Unload Supplier_Info
    Unload SalesForm
    Unload Stocks
    Unload PurchasedItem
    Load HelpForm
End Sub
Private Sub Form_Resize()
    HelpForm.Width = Me.ScaleWidth
    HelpForm.Height = Me.ScaleHeight
    
    
    ResizeControls Me, initialcontrollist
End Sub
