VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SplashScreen 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9975
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   19530
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
   ScaleHeight     =   9975
   ScaleWidth      =   19530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Progress_timer 
      Interval        =   100
      Left            =   17280
      Top             =   7920
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   9000
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   2640
      Width           =   16575
   End
   Begin VB.Image SplashScreen 
      Height          =   9615
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   18975
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private InitialControlList() As ControlInitial
Option Explicit

Private iTBL As TaskbarList

Private bToggle As Boolean
Private hIcoOvr As Long


Private crProgress As Currency
Private Const crProgressMax As Currency = 100















Private Sub Form_Load()
Set iTBL = New TaskbarList
SplashScreen.Move 0, 0, Me.width, Me.height

Progress_timer.Enabled = True
InitialControlList = GetLocation(Me)
ReSizePosForm Me, Me.height, Me.width, Me.Left, Me.Top

End Sub




Private Sub Form_Resize()
SplashScreen.width = Me.ScaleWidth
SplashScreen.height = Me.ScaleHeight
ResizeControls Me, InitialControlList
End Sub

Private Sub progress_timer_Timer()
PBcolor ProgressBar1, vbWhite, vbBlack
Loading.Caption = "Loading Please Wait...." & " " & ProgressBar1.Value & "%"
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = ProgressBar1.Max Then
Progress_timer.Enabled = False
Unload Me
End If
iTBL.SetProgressValue Me.hwnd, crProgress, crProgressMax
crProgress = crProgress + 1

If crProgress = crProgressMax Then
    Progress_timer.Enabled = False
    Unload Me
   
End If
End Sub



