VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RegistrationForm 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "RegistrationForm.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   15000
      Top             =   9960
   End
   Begin VB.TextBox retypepasstxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "."
      TabIndex        =   12
      Top             =   7440
      Width           =   6375
   End
   Begin VB.TextBox passtxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "."
      TabIndex        =   11
      Top             =   6360
      Width           =   6375
   End
   Begin VB.TextBox addresstxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   5280
      Width           =   6375
   End
   Begin VB.TextBox phtxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   4200
      Width           =   6375
   End
   Begin VB.TextBox mailtxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3120
      Width           =   6375
   End
   Begin VB.TextBox nametxt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   6375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   7440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   44114
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   60
      Left            =   15855
      TabIndex        =   20
      Top             =   9900
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   106
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin ShopManagementSystem.LabelPlus lblErrorMessage 
      Height          =   735
      Left            =   15720
      TabIndex        =   21
      Top             =   9840
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
      Caption         =   "RegistrationForm.frx":6783B
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
   Begin ShopManagementSystem.LabelPlus lblAlreadyHaveAnAccount 
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   8400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "RegistrationForm.frx":6785B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin ShopManagementSystem.ucNeumorphism cmdlogin2 
      Height          =   1455
      Left            =   4680
      TabIndex        =   16
      Top             =   7920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
      BackColor       =   8421504
      MousePointer    =   0
   End
   Begin ShopManagementSystem.LabelPlus lblRegister 
      Height          =   495
      Left            =   2160
      TabIndex        =   15
      Top             =   8400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "RegistrationForm.frx":678A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin ShopManagementSystem.ucNeumorphism cmdregister2 
      Height          =   1455
      Left            =   1320
      TabIndex        =   14
      Top             =   7920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
      BackColor       =   8421504
      MousePointer    =   0
   End
   Begin VB.Label Label_Mail 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
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
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label_Retype_Password 
      BackStyle       =   0  'Transparent
      Caption         =   "Retype-Password"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label_Password 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label_Address 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label_Phno 
      BackStyle       =   0  'Transparent
      Caption         =   "Ph No"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1335
      Left            =   -2160
      TabIndex        =   0
      Top             =   240
      Width           =   15615
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   9855
      Left            =   960
      TabIndex        =   18
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   17383
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   2055
      Left            =   -480
      TabIndex        =   19
      Top             =   -240
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   3625
      BackColor       =   4210752
      MousePointer    =   0
   End
   Begin VB.Image RegistrationForm 
      Height          =   10815
      Left            =   0
      Picture         =   "RegistrationForm.frx":678D9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   20895
   End
End
Attribute VB_Name = "RegistrationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial
Option Explicit



Private Sub addresstxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        passtxt.SetFocus
    End If
End Sub

Private Sub cmdlogin2_Click()
Unload Me
LoginForm.Show
End Sub
Private Sub clearcontrol()
nametxt = ""
mailtxt = ""
phtxt = ""
addresstxt = ""
passtxt = ""
retypepasstxt = ""
End Sub




Private Sub cmdregister2_Click()


    If nametxt = "" Then
        If mailtxt = "" Then
            If phtxt = "" Then
                If addresstxt = "" Then
                    If passtxt = "" Then
                        If retypepasstxt = "" Then
                            'MsgBox "Please Fill All The Input Fields", vbOKOnly + vbCritical, App.Title
                            lblErrorMessage.Caption = "Please Fill All The Input Fields...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
   
    If nametxt = "" Then
        'MsgBox "Please Enter The Name", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Please Enter The Name...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
    
    ElseIf mailtxt = "" Then
        'MsgBox "Please Enter The E-mail Address", vbOKOnly + vbCritical, App.Title
         lblErrorMessage.Caption = "Please Enter The E-mail Address...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
    
    ElseIf phtxt = "" Then
        'MsgBox "Please Enter The Phone Number", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Please Enter The Phone Number...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
    
    ElseIf addresstxt = "" Then
        'MsgBox "Please Enter The Address", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Please Enter The Address...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
    
    ElseIf passtxt = "" Then
        'MsgBox "Please Enter The Password", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Please Enter The Password...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
    
    ElseIf retypepasstxt = "" Then
        'MsgBox "Please Confirm Your Password", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Please Confirm Your Password...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
        retypepasstxt.SetFocus
        
    
    ElseIf retypepasstxt <> passtxt Then
        'MsgBox "Password Didn't Matched Please Confirm Your Password", vbOKOnly + vbCritical, App.Title
        lblErrorMessage.Caption = "Password Didn't Matched Please Confirm Your Password...!!!"
                            lblErrorMessage.Visible = True
                            PB1.Visible = True
                            Timer1.Enabled = True
        retypepasstxt.SetFocus
        
        Exit Sub
        Else
             VarRegistrationForm.AddNew

            VarRegistrationForm.Fields(1) = nametxt.text
            VarRegistrationForm.Fields(2) = mailtxt.text
            VarRegistrationForm.Fields(3) = Val(phtxt.text)
            VarRegistrationForm.Fields(4) = addresstxt.text
            VarRegistrationForm.Fields(5) = passtxt.text
            VarRegistrationForm.Fields(6) = retypepasstxt.text
            VarRegistrationForm.Fields(7) = DTPicker1.Value
            VarRegistrationForm.Update
            VarRegistrationForm.MoveFirst
            While VarRegistrationForm.EOF <> True
            VarLoginForm.AddNew
            VarLoginForm.Fields(1) = VarRegistrationForm(1)
            VarLoginForm.Fields(2) = VarRegistrationForm(5)
            VarLoginForm.Update
            VarRegistrationForm.MoveNext
            Wend

            'MsgBox "record saved successfully"
            lblErrorMessage.Caption = "Record saved successfully"
            
            lblErrorMessage.Visible = True
            clearcontrol
    End If




End Sub

Private Sub Form_Load()

                            lblErrorMessage.Visible = False
                            PB1.Visible = False
                            Timer1.Enabled = False

Label_Name.FontName = "Tahoma"
Label_Mail(1).FontName = "Tahoma"
Label_Phno(0).FontName = "Tahoma"
Label_Address(1).FontName = "Tahoma"
Label_Password(0).FontName = "Tahoma"
Label_Retype_Password(0).FontName = "Tahoma"




Label_Name.FontBold = True
Label_Mail(1).FontBold = True
Label_Phno(0).FontBold = True
Label_Address(1).FontBold = True
Label_Password(0).FontBold = True
Label_Retype_Password(0).FontBold = True






RegistrationForm.Move 0, 0, Me.Width, Me.Height
initialcontrollist = GetLocation(Me)
ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top



End Sub
Private Sub Form_Resize()
RegistrationForm.Width = Me.ScaleWidth
RegistrationForm.Height = Me.ScaleHeight
ResizeControls Me, initialcontrollist
End Sub






Private Sub lblAlreadyHaveAnAccount_Click()
Call cmdlogin2_Click
End Sub

Private Sub lblRegister_Click()
Call cmdregister2_Click
End Sub

Private Sub mailtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        phtxt.SetFocus
    End If
End Sub

Private Sub nametxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mailtxt.SetFocus
    End If
End Sub





Private Sub passtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        retypepasstxt.SetFocus
    End If
End Sub

Private Sub phtxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        addresstxt.SetFocus
    End If
End Sub



Private Sub retypepasstxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdregister2.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
 PBcolor PB1, &H404040, &HC0C0C0
    If PB1.Value = PB1.Min Then
        PB1.Value = PB1.Max
    End If
    lblRegister.Enabled = False
            cmdregister2.Enabled = False
            lblAlreadyHaveAnAccount.Enabled = False
            cmdlogin2.Enabled = False
            nametxt.Enabled = False
            
mailtxt.Enabled = False

phtxt.Enabled = False

addresstxt.Enabled = False

passtxt.Enabled = False
retypepasstxt.Enabled = False

            
            
            
     PB1.Value = PB1.Value - 1
     
     
        
        If PB1.Value = PB1.Min Then
            Timer1.Enabled = False
            lblErrorMessage.Visible = False
            PB1.Visible = False
           
           lblRegister.Enabled = True
            cmdregister2.Enabled = True
            lblAlreadyHaveAnAccount.Enabled = True
            cmdlogin2.Enabled = True
            nametxt.Enabled = True
            
            mailtxt.Enabled = True

            phtxt.Enabled = True

            addresstxt.Enabled = True

            passtxt.Enabled = True
            retypepasstxt.Enabled = True
            
            End If
End Sub
