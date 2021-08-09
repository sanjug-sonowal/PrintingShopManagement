VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SplashScreen 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9630
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   19170
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   19170
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   9000
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Progress_timer 
      Interval        =   100
      Left            =   17280
      Top             =   7920
   End
   Begin VB.Label Loading 
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
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   8520
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PRINTING SHOP MANAGEMENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1575
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   16575
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   2895
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   5106
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Image Form1 
      Height          =   9735
      Left            =   0
      Picture         =   "SplashScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19215
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private initialcontrollist() As ControlInitial
Option Explicit

Private MyProgressState As TaskbarList
Private crProgress As Currency
Private Const crProgressMax As Currency = 100

Private Sub Form_Load()
Set MyProgressState = New TaskbarList
    SplashScreen.Move 0, 0, Me.Width, Me.Height

    Progress_timer.Enabled = True
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub




Private Sub Form_Resize()
    SplashScreen.Width = Me.ScaleWidth
    SplashScreen.Height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub

Private Sub progress_timer_Timer()
    PBcolor ProgressBar1, vbWhite, vbBlack
    Loading.Caption = "Loading Please Wait...." & " " & ProgressBar1.Value & "%"
    ProgressBar1.Value = ProgressBar1.Value + 1
        
        If ProgressBar1.Value = ProgressBar1.Max Then
            Progress_timer.Enabled = False
            Unload Me
        End If
            MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
            crProgress = crProgress + 1

        If crProgress = crProgressMax Then
            Progress_timer.Enabled = False
            Unload Me
            LoginForm.Show
   
        End If
End Sub



