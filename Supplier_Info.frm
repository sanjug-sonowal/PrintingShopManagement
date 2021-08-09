VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Supplier_Info 
   BorderStyle     =   0  'None
   Caption         =   "Supplier Information"
   ClientHeight    =   11415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11415
   ScaleWidth      =   20835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR RECORDS"
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE RECORDS"
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HOME"
      Height          =   495
      Left            =   17760
      TabIndex        =   14
      Top             =   10680
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT REPORTS"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE RECORDS"
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "DISPLAY RECORDS"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Frame Supplier_Info_Frame 
      BackColor       =   &H00808080&
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   20535
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Caption         =   "SUPPLIER DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   9015
         Left            =   6720
         TabIndex        =   15
         Top             =   480
         Width           =   13695
         Begin MSDataGridLib.DataGrid Supplier_Details_DataGrid 
            Height          =   8535
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   15055
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
      End
      Begin VB.TextBox Supplier_Address_txt 
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox Supplier_Mob_No_txt 
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Supplier_Name_txt 
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Supplier_Id_txt 
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Invoice_No_txt 
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Top             =   1560
         Width           =   3015
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism2 
         Height          =   4575
         Left            =   -240
         TabIndex        =   22
         Top             =   5280
         Width           =   7335
         _extentx        =   12938
         _extenty        =   8070
         backcolor       =   4210752
         mousepointer    =   0
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   5
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER MOB NO"
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
         Left            =   480
         TabIndex        =   4
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   3
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   2
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INCOICE ID"
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
         Left            =   480
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
      End
      Begin ShopManagementSystem.ucNeumorphism ucNeumorphism1 
         Height          =   4935
         Left            =   -240
         TabIndex        =   21
         Top             =   840
         Width           =   7335
         _extentx        =   12938
         _extenty        =   8705
         backcolor       =   4210752
         mousepointer    =   0
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER DETAILS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   5400
      TabIndex        =   20
      Top             =   120
      Width           =   9135
   End
   Begin VB.Image Supplier_Info_Background_Image 
      Height          =   11295
      Left            =   0
      Picture         =   "Supplier_Info.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20775
   End
End
Attribute VB_Name = "Supplier_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialcontrollist() As ControlInitial






Private Sub cmdClear_Click()
clearcontrol
End Sub

Private Sub CmdDelete_Click()
Dim invoice As Integer
invoice = InputBox("Enter Invoice Number Which You Want to Delete From Database")

'---------------------------------------------------------------------------------------------------------------------
'DELETE FROM PURCHASED DETAILS TABLE IN THE DATABASE
'---------------------------------------------------------------------------------------------------------------------

VarRecordDelete.Open "select * from SupplierDetails where Invoice_No=" & invoice, cn, adOpenDynamic, adLockOptimistic
If Not (VarRecordDelete.EOF) Then
VarRecordDelete.Delete
End If

VarRecordDelete.Close
Call cmdDisplay_Click

End Sub

Private Sub cmdDisplay_Click()
 Dim VarSupp As New Recordset
 VarSupp.Open "select * from SupplierDetails", cn, adOpenDynamic, adLockOptimistic
Set Supplier_Details_DataGrid.DataSource = VarSupp
Supplier_Details_DataGrid.Refresh
Set VarSupp = Nothing

End Sub

Private Sub cmdExit_Click()
End
DailyLogSheet.Show
End Sub

Private Sub cmdPrint_Click()
SupplierReport.Show
End Sub

Private Sub cmdUpdate_Click()
VarInvoiceUpdate.Open "Select * from SupplierDetails where Invoice_No=" & Val(Invoice_No_txt.text) & "", cn, adOpenDynamic, adLockOptimistic
If VarInvoiceUpdate.EOF Then
MsgBox "Data Not Found"
VarInvoiceUpdate.Close
Else
VarInvoiceUpdate(0) = Invoice_No_txt.text
VarInvoiceUpdate(1) = Supplier_Id_txt.text
VarInvoiceUpdate(2) = Supplier_Name_txt.text
VarInvoiceUpdate(3) = Supplier_Mob_No_txt.text
VarInvoiceUpdate(4) = Supplier_Address_txt.text
VarInvoiceUpdate.Update



MsgBox "Updated Successfully"

Call cmdDisplay_Click


VarInvoiceUpdate.Close
End If
clearcontrol
End Sub
Private Function clearcontrol()
Invoice_No_txt.text = ""
Supplier_Id_txt.text = ""
Supplier_Name_txt.text = ""
Supplier_Mob_No_txt.text = ""
Supplier_Address_txt.text = ""
End Function

Private Sub Command4_Click()
DailyLogSheet.Show
End Sub

Private Sub Form_Load()




Supplier_Info_Background_Image.Move 0, 0, Me.Width, Me.Height

initialcontrollist = GetLocation(Me)
ReSizePosForm Me, Me.Height, Me.Width, Me.Left, Me.Top

End Sub

Private Sub Form_Resize()
    Supplier_Info_Background_Image.Width = Me.ScaleWidth
    Supplier_Info_Background_Image.Height = Me.ScaleHeight
    
    ResizeControls Me, initialcontrollist
End Sub

Private Sub Supplier_Details_DataGrid_Click()
'-----------------------------------------------------------------------------------------
'Variable's For Displaying Selected Records from DataGrid
'-----------------------------------------------------------------------------------------
 
 Dim Invoice_No As Integer
 Dim Supplier_Id As Integer
 Dim supplier_name As String
 Dim Supplier_Mob_No As String
 Dim Supplier_Address As String

 
 '----------------------------------------------------------------------------------------
 'Fetching Data Into The Textbox From the Database
 '----------------------------------------------------------------------------------------
 
 Invoice_No = Supplier_Details_DataGrid.Columns(0)
 Supplier_Id = Supplier_Details_DataGrid.Columns(1)
 supplier_name = Supplier_Details_DataGrid.Columns(2)
 Supplier_Mob_No = Supplier_Details_DataGrid.Columns(3)
 Supplier_Address = Supplier_Details_DataGrid.Columns(4)
 
 
 '----------------------------------------------------------------------------------------
 'Displaying Selected Records from DataGrid Into the Textbox
 '----------------------------------------------------------------------------------------
 
 Invoice_No_txt.text = Invoice_No
 Supplier_Id_txt.text = Supplier_Id
 Supplier_Name_txt.text = supplier_name
 Supplier_Mob_No_txt.text = Supplier_Mob_No
 Supplier_Address_txt.text = Supplier_Address
 

'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
End Sub
