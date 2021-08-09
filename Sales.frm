VERSION 5.00
Begin VB.Form Sales 
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
   Begin VB.Image Sales_Background_Image 
      Height          =   10335
      Left            =   0
      Picture         =   "Sales.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20775
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Private Sub Form_Load()


Sales_Background_Image.Move 0, 0, Me.width, Me.height

initialcontrollist = GetLocation(Me)
ReSizePosForm Me, Me.height, Me.width, Me.Left, Me.Top

End Sub

Private Sub Form_Resize()
    Sales_Background_Image.width = Me.ScaleWidth
    Sales_Background_Image.height = Me.ScaleHeight
    
    ResizeControls Me, initialcontrollist
End Sub
