VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ChangePassword 
   Caption         =   "ForgotPassword"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   15000
      Top             =   10080
   End
   Begin VB.Timer Progress_State 
      Left            =   16320
      Top             =   5040
   End
   Begin VB.TextBox confirmpasstxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   8400
      PasswordChar    =   "."
      TabIndex        =   3
      Top             =   8400
      Width           =   6375
   End
   Begin VB.TextBox newpasstxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   8400
      PasswordChar    =   "."
      TabIndex        =   2
      Top             =   7080
      Width           =   6375
   End
   Begin VB.TextBox verifymailtxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   1
      Top             =   3840
      Width           =   6375
   End
   Begin VB.TextBox checkusertxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   0
      Top             =   2040
      Width           =   6375
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   60
      Left            =   9375
      TabIndex        =   19
      Top             =   5460
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   106
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin ShopManagementSystem.LabelPlus lblErrorMessage 
      Height          =   735
      Left            =   9240
      TabIndex        =   20
      Top             =   5400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      BackColor       =   4210752
      BorderColor     =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":0000
      CaptionPaddingY =   10
      CaptionShadow   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowSize      =   3
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus Label5 
      Height          =   735
      Left            =   8400
      TabIndex        =   18
      Top             =   4560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1296
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":0020
      CaptionShadow   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus Label4 
      Height          =   735
      Left            =   8400
      TabIndex        =   17
      Top             =   2640
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1296
      BackColor       =   4210752
      BorderCornerLeftTop=   7
      BorderCornerRightTop=   7
      BorderCornerBottomRight=   7
      BorderCornerBottomLeft=   7
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":0040
      CaptionShadow   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowSize      =   7
      ShadowOffsetY   =   5
      ShadowColorOpacity=   30
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblChangePassword 
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   9600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":0060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ChangePassword_Btn 
      Height          =   1575
      Left            =   8040
      TabIndex        =   14
      Top             =   9000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.LabelPlus lblVerify 
      Height          =   375
      Left            =   16560
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":009E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.LabelPlus lblCheck 
      Height          =   375
      Left            =   16560
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "ChangePassword.frx":00CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin ShopManagementSystem.ucNeumorphism Verify_btn 
      Height          =   1575
      Left            =   15480
      TabIndex        =   11
      Top             =   3360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2778
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism Check_Btn 
      Height          =   1575
      Left            =   15480
      TabIndex        =   10
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2778
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FORGOT PASSWORD"
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
      Height          =   1335
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   12135
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   10200
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   3015
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   10215
      Left            =   7560
      TabIndex        =   16
      Top             =   960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   18018
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Image ChangePassword 
      Height          =   11295
      Left            =   0
      Picture         =   "ChangePassword.frx":00F4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20655
   End
End
Attribute VB_Name = "ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------------------------------------------'
'This Code is for Making responsive form

    Private initialcontrollist() As ControlInitial
'------------------------------------------------------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------------------------------------------------------'
'This Code is for My Taskbar Progress state

    Private MyProgressState As TaskbarList
    Private crProgress As Currency
    Private Const crProgressMax As Currency = 100
'------------------------------------------------------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------------------------------------------------------'
'This Code is for My Password checking if there are duplicate username

    Dim start As Integer
'------------------------------------------------------------------------------------------------------------------------------------'

'------------------------------------------------------------------------------------------------------------------------------------'

Private Sub ChangePassword_Btn_Click()
Progress_State.Interval = 100
crProgress = 0
    If newpasstxt.text = confirmpasstxt.text Then
        VarLoginForm.Fields("Password") = confirmpasstxt.text
        VarRegistrationForm.Fields("Password") = confirmpasstxt
        VarRegistrationForm.Fields("RetypePassword") = confirmpasstxt
        MyProgressState.SetProgressState Me.hwnd, TBPF_INDETERMINATE
        Progress_State.Enabled = True
        VarLoginForm.Update
        VarRegistrationForm.Update
            'MsgBox "Password Changed Successfully", vbInformation, "Password Change: Success"
            Timer1.Enabled = True
            PB1.Visible = True
            lblErrorMessage.Visible = True
            
            
             MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
                crProgress = crProgressMax
                    If crProgress = crProgressMax Then
                        Progress_State.Enabled = False
                    End If
                        
                        
    Else
        MyProgressState.SetProgressState Me.hwnd, TBPF_INDETERMINATE
        'MsgBox "Password Does not matched,Please Enter Correct Details", vbExclamation, "Change Password:Failed"
        Timer1.Enabled = True
            lblErrorMessage.Caption = "Password Does not matched,Please Enter Correct Details...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = False
        newpasstxt.text = ""
        confirmpasstxt.text = ""
    End If
End Sub
'---------------------------------------------------------------------------------------------------------------------------------'






'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub Check_Btn_Click()
Progress_State.Interval = 100
crProgress = 0
     VarLoginForm.MoveFirst
        VarLoginForm.Find "UserName='" & checkusertxt & "'", 0, adSearchForward, start
            If Not VarLoginForm.EOF Then
                'correctuser
                    If checkusertxt = VarLoginForm.Fields!UserName Then
                         Label4.Caption = "UserName Found in the database"
                         Label4.ForeColor = &H8000&
                         Label5.Visible = False
                         
                         Label4.Visible = True
                         Label7.Visible = True
                         verifymailtxt.Visible = True
                         verifymailtxt.SetFocus
                         MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
                         crProgress = crProgressMax
                         
                                If crProgress = crProgressMax Then
                                    Progress_State.Enabled = False
                                End If
                                
                                Else
                                Label4.Visible = True
                                    Label4.Caption = "UserName Not Found ..Sorry Can't ReSet the password!!!! "
                                    Label4.ForeColor = &HFFFFFF
                                    Label5.Visible = False
                                    Label7.Visible = False
                                    verifymailtxt.Visible = False
                                    
                                    checkusertxt.SetFocus
                                    MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
                                    Progress_State.Enabled = True
                    End If
            Else
                Label4.Visible = True
                Label5.Visible = False
                Label4.Caption = "UserName Not Found ..Sorry Can't ReSet the password!!!! "
                Label4.ForeColor = &HFFFFFF
                Label7.Visible = False
                                    verifymailtxt.Visible = False
                checkusertxt.SetFocus
                MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
                Progress_State.Enabled = True
  
           End If
End Sub
'--------------------------------------------------------------------------------------------------------------------------------'



Private Sub Check_Btn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        verifymailtxt.SetFocus
    End If
End Sub

Private Sub checkusertxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Check_Btn.SetFocus
    End If
End Sub



Private Sub confirmpasstxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ChangePassword_Btn.SetFocus
    End If
End Sub

'--------------------------------------------------------------------------------------------------------------------------------'

Private Sub Form_Load()

Timer1.Enabled = False
            lblErrorMessage.Visible = False
            PB1.Visible = False
Set MyProgressState = New TaskbarList
start = 1

    Label9.FontName = "Tahoma"
    Label9.fontsize = 10
    Label9.BackStyle = 0
    Label9.ForeColor = vbWhite
    Label9.FontBold = True
    
   
    
    
    Label10.FontName = "Tahoma"
    Label10.fontsize = 10
    Label10.BackStyle = 0
    Label10.ForeColor = vbWhite
    Label10.FontBold = True
    

    checkusertxt.FontName = "Tahoma"
    checkusertxt.fontsize = 10

    verifymailtxt.FontName = "Tahoma"
    verifymailtxt.fontsize = 10

    newpasstxt.FontName = "Tahoma"
    newpasstxt.fontsize = 10

    confirmpasstxt.FontName = "Tahoma"
    confirmpasstxt.fontsize = 10


    Label4.Visible = False
    Label5.Visible = False
    
    Label7.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    lblChangePassword.Visible = False


    newpasstxt.Visible = False
    confirmpasstxt.Visible = False

    verifymailtxt.Visible = False
    ChangePassword_Btn.Visible = False

    

    Label8.FontName = "Tahoma"
    Label8.fontsize = 12

   

    Label3.FontName = "Tahoma"
    Label3.fontsize = 48

   


    ChangePassword.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top
End Sub
'---------------------------------------------------------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub Form_Resize()
    ChangePassword.Width = Me.ScaleWidth
    ChangePassword.Height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub
'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub Form_Unload(Cancel As Integer)
    Set MyProgressState = Nothing
End Sub
'----------------------------------------------------------------------------------------------------------------------------------'






Private Sub lblChangePassword_Click()
Call ChangePassword_Btn_Click
End Sub

Private Sub lblCheck_Click()
Call Check_Btn_Click
End Sub

Private Sub lblVerify_Click()
Call Verify_btn_Click
End Sub

Private Sub newpasstxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        confirmpasstxt.SetFocus
    End If
End Sub

'----------------------------------------------------------------------------------------------------------------------------------'

Private Sub Progress_State_Timer()
    MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
    crProgress = crProgressMax
End Sub
'----------------------------------------------------------------------------------------------------------------------------------'


Private Sub Timer1_Timer()
PBcolor PB1, &H404040, &HC0C0C0
    If PB1.Value = PB1.Min Then
        PB1.Value = PB1.Max
    End If
            
     lblErrorMessage.Visible = True
     lblErrorMessage.Caption = "Password Changed Successfully...!!!"
            
     PB1.Value = PB1.Value - 1
     
     
        
        If PB1.Value = PB1.Min Then
            Timer1.Enabled = False
            lblErrorMessage.Visible = False
            PB1.Visible = False
           Unload Me
           LoginForm.Show
           
            
           
           
            
            
            End If
End Sub

'----------------------------------------------------------------------------------------------------------------------------------'
Private Sub Verify_btn_Click()
Progress_State.Interval = 100
crProgress = 0

  VarRegistrationForm.MoveFirst
        Dim str As String
            str = StrComp(VarRegistrationForm.Fields("EMail").Value, verifymailtxt.text, vbTextCompare)
                If str = True Then
                    Label5.Caption = "Account not verified , Can't reset the password"
                    checkusertxt.text = ""
                    checkusertxt.SetFocus
                    verifymailtxt.text = ""
                    Label7.Visible = False
                    
                    Label5.Caption = "Please Enter An Valid Credential"
                    Label5.ForeColor = &HFF&
                    Label5.Visible = True
                    Label7.Visible = False
                    verifymailtxt.Visible = False
                    checkusertxt.SetFocus
                    
                    
                    Label4.Visible = False
                    Label4.ForeColor = &HFF&
                    
                    Label8.ForeColor = &HFF&
                    newpasstxt.Visible = False
                    confirmpasstxt.Visible = False
                    
                    Label9.Visible = False
                    Label10.Visible = False
                    lblChangePassword.Visible = False
                    
                    
                    ChangePassword_Btn.Visible = False
                    MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
                    Progress_State.Enabled = True

            Else
            
                If str = False And checkusertxt = VarLoginForm.Fields!UserName Then
                    Label5.ForeColor = &H8000&
                    Label8.ForeColor = &H8000&
                    Label4.ForeColor = &H8000&
                    Label4.Caption = "Congratulations !!"
                    Label5.Caption = "Account is verified Now,Set your new Password"
                    Label4.Visible = True
                    Label5.Visible = True
                    
                    newpasstxt.Visible = True
                    confirmpasstxt.Visible = True
                    
                    Label9.Visible = True
                    Label10.Visible = True
                    lblChangePassword.Visible = True
                    
                    newpasstxt.SetFocus
                    ChangePassword_Btn.Visible = True
                    MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
                    crProgress = crProgressMax
                        
                        If crProgress = crProgressMax Then
                            Progress_State.Enabled = False
                        End If
                End If



    
        End If

End Sub

Private Sub Verify_btn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        newpasstxt.SetFocus
    End If
End Sub



Private Sub verifymailtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Verify_btn.SetFocus
    End If
End Sub
