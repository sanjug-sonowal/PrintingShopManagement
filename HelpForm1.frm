VERSION 5.00
Begin VB.Form HelpForm1 
   Caption         =   "Help Form"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21375
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   21375
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   10335
      Left            =   0
      Picture         =   "HelpForm1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21255
   End
End
Attribute VB_Name = "HelpForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------'
'This Code is for making responsive form

Private initialcontrollist() As ControlInitial
'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub Form_Load()
 HelpForm1.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub
Private Sub Form_Resize()
    HelpForm1.Width = Me.ScaleWidth
    HelpForm1.Height = Me.ScaleHeight
    
    
    ResizeControls Me, initialcontrollist
End Sub
