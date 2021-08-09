VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginForm 
   BackColor       =   &H00808080&
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox usertxt1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   2400
      Width           =   6855
   End
   Begin VB.TextBox passtxt1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3720
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14760
      Top             =   9720
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   60
      Left            =   15615
      TabIndex        =   4
      Top             =   9660
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   106
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Progress_State 
      Interval        =   250
      Left            =   4800
      Top             =   9360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2160
      Top             =   8040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"LoginForm.frx":0000
      OLEDBString     =   $"LoginForm.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ShopManagementSystem.LabelPlus lblErrorMessage 
      Height          =   735
      Left            =   15480
      TabIndex        =   8
      Top             =   9600
      Width           =   4455
      _extentx        =   7858
      _extenty        =   1296
      backcolor       =   4210752
      bordercolor     =   4210752
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      captionalignmenth=   1
      captionalignmentv=   1
      caption         =   "LoginForm.frx":013C
      captionpaddingy =   10
      captionshadow   =   -1  'True
      font            =   "LoginForm.frx":015C
      forecolor       =   16777215
      shadowsize      =   3
      shadowoffsety   =   5
      shadowcoloropacity=   30
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "LoginForm.frx":0188
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN FORM"
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
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin ShopManagementSystem.LabelPlus lblForgotPassword 
      Height          =   855
      Left            =   3960
      TabIndex        =   12
      Top             =   5880
      Width           =   2175
      _extentx        =   3836
      _extenty        =   1508
      backcoloropacity=   0
      backshadow      =   0   'False
      bordercolor     =   16777215
      captionalignmenth=   1
      captionalignmentv=   1
      caption         =   "LoginForm.frx":01B4
      captionbordercolor=   16777215
      font            =   "LoginForm.frx":01F2
      forecolor       =   16777215
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "LoginForm.frx":021A
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin ShopManagementSystem.ucNeumorphism cmdforgotpassword 
      Height          =   1575
      Left            =   3360
      TabIndex        =   11
      Top             =   5520
      Width           =   3495
      _extentx        =   6165
      _extenty        =   2778
      mousepointer    =   0
   End
   Begin ShopManagementSystem.LabelPlus lblNotAnMemberYet 
      Height          =   855
      Left            =   5880
      TabIndex        =   10
      Top             =   4680
      Width           =   2535
      _extentx        =   4471
      _extenty        =   1508
      backcoloropacity=   0
      backshadow      =   0   'False
      bordercolor     =   16777215
      captionalignmenth=   1
      captionalignmentv=   1
      caption         =   "LoginForm.frx":0246
      captionbordercolor=   16777215
      font            =   "LoginForm.frx":0288
      forecolor       =   16777215
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "LoginForm.frx":02B0
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin ShopManagementSystem.ucNeumorphism cmdregister1 
      Height          =   1575
      Left            =   5280
      TabIndex        =   9
      Top             =   4320
      Width           =   3735
      _extentx        =   6588
      _extenty        =   2778
      mousepointer    =   0
   End
   Begin ShopManagementSystem.LabelPlus lblLogin 
      Height          =   855
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1508
      backcoloropacity=   0
      backshadow      =   0   'False
      bordercolor     =   16777215
      captionalignmenth=   1
      captionalignmentv=   1
      caption         =   "LoginForm.frx":02DC
      captionbordercolor=   16777215
      font            =   "LoginForm.frx":0306
      forecolor       =   16777215
      shadowcoloropacity=   0
      calloutalign    =   0
      calloutwidth    =   0
      calloutlen      =   0
      mousepointer    =   0
      iconfont        =   "LoginForm.frx":032E
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin ShopManagementSystem.ucNeumorphism cmdlogin1 
      Height          =   1575
      Left            =   1440
      TabIndex        =   5
      Top             =   4320
      Width           =   3495
      _extentx        =   6165
      _extenty        =   2778
      mousepointer    =   0
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
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
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
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
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
      Height          =   7455
      Left            =   960
      TabIndex        =   13
      Top             =   1080
      Width           =   8535
      _extentx        =   15055
      _extenty        =   13150
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism3 
      Height          =   2175
      Left            =   1320
      TabIndex        =   14
      Top             =   -240
      Width           =   7695
      _extentx        =   13573
      _extenty        =   3836
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
      Height          =   11655
      Left            =   -360
      TabIndex        =   15
      Top             =   -360
      Width           =   20895
      _extentx        =   36856
      _extenty        =   20558
      backcolor       =   4210752
      mousepointer    =   0
   End
   Begin VB.Image LoginForm 
      Height          =   10935
      Left            =   -600
      Picture         =   "LoginForm.frx":035A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21375
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------'
'This Code is for making responsive form

Private initialcontrollist() As ControlInitial
'---------------------------------------------------------------------------------------------------------------------------------'
'This Code is for My Taskbar Progress state

    Private MyProgressState As TaskbarList
    Private crProgress As Currency
    Private Const crProgressMax As Currency = 100
'---------------------------------------------------------------------------------------------------------------------------------'
'This Code is for checking password while the username is duplicate
Dim start As Integer
'---------------------------------------------------------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub cmdforgotpassword_Click()
   Unload Me
   ChangePassword.Show
End Sub
'---------------------------------------------------------------------------------------------------------------------------------'

'---------------------------------------------------------------------------------------------------------------------------------'

Private Sub cmdlogin1_Click()
Progress_State.Interval = 100
crProgress = 0
    Progress_State.Enabled = True
    '---------------------------------------------------------------------------------
   
    '---------------------------------------------------------------------------------
   
    
    If usertxt1 = "" Then
        If passtxt1 = "" Then
            MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
            'MsgBox "Username and Password Is Empty", vbOKOnly + vbCritical, App.Title
            
            lblErrorMessage.Caption = "Username and Password Is Empty...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = True
            Timer1.Enabled = True
            usertxt1.SetFocus
            
            '--------------------------------------------
      

        '----------------------------------------------------------------
            
            
            
            
                
            
            Exit Sub
        End If
    End If
        If passtxt1 = "" Then
            MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
             lblErrorMessage.Caption = "Sorry Your Password Is Empty...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = True
            Timer1.Enabled = True
            passtxt1.SetFocus
           
            
            
            Exit Sub
        End If
            If usertxt1 = "" Then
                MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
                
            lblErrorMessage.Caption = "Sorry Your Username Is Empty...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = True
            Timer1.Enabled = True
            usertxt1.SetFocus
                
                Exit Sub
            End If
   

'Login Validation Code
VarLoginForm.MoveFirst
  
VarLoginForm.Find "UserName='" & usertxt1 & "'", 0, adSearchForward, start
    If Not VarLoginForm.EOF Then
       'correctuser
         If usertxt1 = VarLoginForm.Fields!UserName Then
       
            If passtxt1 = VarLoginForm.Fields!password Then
                'correct password
                MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
                         crProgress = crProgressMax
                         
                                If crProgress = crProgressMax Then
                                    Progress_State.Enabled = False
                                End If
                Unload Me
                MainForm.Show
                DailyLogSheet.Show
                PurchasedItem.Visible = False
                Stocks.Visible = False
                Supplier_Info.Visible = False
                SalesForm.Visible = False
                
                
           Else
                MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
                'MsgBox "Please Enter An Valid Username And Password", vbOKOnly + vbCritical, App.Title
                lblErrorMessage.Caption = "Please Enter An Valid Username And Password...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = True
            Timer1.Enabled = True
            
            End If
        End If
Else
           
         ' incorrect password
          start = VarLoginForm.AbsolutePosition + 1
          
          If start = VarLoginForm.RecordCount Then
             MyProgressState.SetProgressState Me.hwnd, TBPF_ERROR
             
             'MsgBox "Incorrect Username or Password", vbOKOnly + vbCritical, App.Title
             
             lblErrorMessage.Caption = "Incorrect Username or Password...!!!"
            lblErrorMessage.Visible = True
            PB1.Visible = True
            Timer1.Enabled = True
             Exit Sub
          End If
         
         Call cmdlogin1_Click
       
       End If
  
     
End Sub
'------------------------------------------------------------------------------------------------------------------------------'



'------------------------------------------------------------------------------------------------------------------------------'
Private Sub cmdregister1_Click()
    Unload Me
    RegistrationForm.Show
End Sub
'------------------------------------------------------------------------------------------------------------------------------'





Private Sub Command1_Click()
SupplierReport.Show
End Sub

Private Sub Command2_Click()
LogSheetReport.Show
End Sub

Private Sub Command3_Click()
StockReport.Show
End Sub

Private Sub Command4_Click()
SupplierReport.Show
End Sub

'------------------------------------------------------------------------------------------------------------------------------'
Private Sub Form_Load()
    
    
    Set MyProgressState = New TaskbarList
    Progress_State.Enabled = False
    lblErrorMessage.Visible = False
    
    PB1.Visible = False
   
    Timer1.Enabled = False
    
    
    
    Adodc1.Visible = False
    start = 1
    Me.WindowState = 2
    


    


    passtxt1.FontName = "Tahoma"
    passtxt1.fontsize = 10


    


    LoginForm.Move 0, 0, Me.Width, Me.Height
    initialcontrollist = GetLocation(Me)
    ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub
'--------------------------------------------------------------------------------------------------------------------------------'

'--------------------------------------------------------------------------------------------------------------------------------'
Private Sub Form_Resize()
    LoginForm.Width = Me.ScaleWidth
    LoginForm.Height = Me.ScaleHeight
    ResizeControls Me, initialcontrollist
End Sub








Private Sub lblForgotPassword_Click()
Call cmdforgotpassword_Click
End Sub

Private Sub lblLogin_Click()
Call cmdlogin1_Click
End Sub

Private Sub lblNotAnMemberYet_Click()
Call cmdregister1_Click
End Sub

Private Sub passtxt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdlogin1.SetFocus
    End If
End Sub



Private Sub Progress_State_Timer()
    MyProgressState.SetProgressValue Me.hwnd, crProgress, crProgressMax
    crProgress = crProgressMax
End Sub



Private Sub Timer1_Timer()

    PBcolor PB1, &H404040, &HC0C0C0
    If PB1.Value = PB1.Min Then
        PB1.Value = PB1.Max
    End If
    cmdlogin1.Enabled = False
            lblLogin.Enabled = False
            lblNotAnMemberYet.Enabled = False
            cmdregister1.Enabled = False
            lblForgotPassword.Enabled = False
            cmdforgotpassword.Enabled = False
            usertxt1.Enabled = False
            passtxt1.Enabled = False
            
            
            
     PB1.Value = PB1.Value - 1
     
     
        
        If PB1.Value = PB1.Min Then
            Timer1.Enabled = False
            lblErrorMessage.Visible = False
            PB1.Visible = False
            cmdlogin1.Enabled = True
            lblLogin.Enabled = True
            lblNotAnMemberYet.Enabled = True
            cmdregister1.Enabled = True
            lblForgotPassword.Enabled = True
            cmdforgotpassword.Enabled = True
            usertxt1.Enabled = True
            passtxt1.Enabled = True
            
            
            End If
            
        
End Sub

'--------------------------------------------------------------------------------------------------------------------------------'



Private Sub usertxt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        passtxt1.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyProgressState = Nothing
End Sub










