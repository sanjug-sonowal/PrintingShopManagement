Attribute VB_Name = "Sanjug_Module5_For_Sub_Main"
Option Explicit

Public cn As ADODB.Connection
Public VarLoginForm As ADODB.Recordset
Public VarLogSheet As ADODB.Recordset
Public VarRegistrationForm As ADODB.Recordset
Public VarStockDetails As ADODB.Recordset
Public VarSupplierDetails As ADODB.Recordset
Public VarPurchasedDetails As ADODB.Recordset
Public VarRecordDelete As ADODB.Recordset

Public VarInvoiceUpdate As ADODB.Recordset
Public VarProductIdUpdate As ADODB.Recordset
Public VarProductNameUpdate As ADODB.Recordset
Public VarSupplierIdUpdate As ADODB.Recordset
Public VarSupplierNameUpdate As ADODB.Recordset
Public VarSupplierMobNoUpdate As ADODB.Recordset
Public VarSupplierAddressUpdate As ADODB.Recordset
Public VarCategoryUpdate As ADODB.Recordset
Public VarBrandUpdate As ADODB.Recordset
Public VarDescriptionUpdate As ADODB.Recordset
Public VarPaperWeightUpdate As ADODB.Recordset
Public VarQuantityUpdate As ADODB.Recordset
Public VarPriceUpdate As ADODB.Recordset
Public VarDateUpdate As ADODB.Recordset

Public Prev As New ADODB.Recordset
Public Dateprev As New ADODB.Recordset










Public Sub SendKeys(text$, Optional wait As Boolean = False)
Dim WshShell As Object
Set WshShell = CreateObject("Wscript.shell")
WshShell.SendKeys text, wait
Set WshShell = Nothing

End Sub










Private Sub main()
'Connection String Code
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ShopManagementSystem_Database.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

'To Initialize Recordset
Set VarLoginForm = New ADODB.Recordset
Set VarLogSheet = New ADODB.Recordset
Set VarRegistrationForm = New ADODB.Recordset
Set VarStockDetails = New ADODB.Recordset
Set VarSupplierDetails = New ADODB.Recordset
Set VarPurchasedDetails = New ADODB.Recordset
Set VarRecordDelete = New ADODB.Recordset

Set VarInvoiceUpdate = New ADODB.Recordset
Set VarProductIdUpdate = New ADODB.Recordset
Set VarProductNameUpdate = New ADODB.Recordset
Set VarSupplierIdUpdate = New ADODB.Recordset
Set VarSupplierNameUpdate = New ADODB.Recordset
Set VarSupplierMobNoUpdate = New ADODB.Recordset
Set VarSupplierAddressUpdate = New ADODB.Recordset
Set VarCategoryUpdate = New ADODB.Recordset
Set VarBrandUpdate = New ADODB.Recordset
Set VarDescriptionUpdate = New ADODB.Recordset
Set VarPaperWeightUpdate = New ADODB.Recordset
Set VarQuantityUpdate = New ADODB.Recordset
Set VarPriceUpdate = New ADODB.Recordset
Set VarDateUpdate = New ADODB.Recordset







VarLoginForm.Open "select * from LoginForm order by Username", cn, adOpenDynamic, adLockOptimistic
VarLogSheet.Open "select * from LogSheet", cn, adOpenDynamic, adLockOptimistic
VarRegistrationForm.Open "select * from RegistrationForm", cn, adOpenDynamic, adLockOptimistic
VarStockDetails.Open "select * from StockDetails", cn, adOpenDynamic, adLockOptimistic
VarSupplierDetails.Open "select * from SupplierDetails", cn, adOpenDynamic, adLockOptimistic
VarPurchasedDetails.Open "select * from PurchasedDetails", cn, adOpenDynamic, adLockOptimistic
VarProductIdUpdate.Open "select * from PurchasedDetails", cn, adOpenDynamic, adLockOptimistic






SplashScreen.Show


End Sub
