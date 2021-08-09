Attribute VB_Name = "mIID"
Option Explicit

'mIID.bas by fafalone
'Revision 16

'Contains UUIDs for working with shell interfaces
'This is a companion module to oleexp, and is not an exhaustive list
'of interface IIDs, just the ones I've used while working with oleexp
'interfaces. These can be used directly in calls as an riid; no need
'for CLSIDFromString; for example SHCreateItemFromIDList(pidl, IID_IShellItem2, isi)

'Note that when you compile a project with the module, only the IIDs that
'are actually used will be compiled in, not the whole module.
'
'Revision History:
'
'Rev. 5
'Added all remaining BHID_ GUID's for IShellItem.BindToHandler

'Rev. 6
'Added UUID_NULL

'Rev. 7
'Added API declare for IsEqualIID

'Rev. 8
'A number of missing IIDs were added; a small error in the automatic conversion script

'Rev. 9
'Major IID additions for oleexp 4.0

'Rev. 10
'IID additions for oleexp 4.2
'GUIDToString function added since the API doesn't seem to work

'Rev. 11
'Fixed IsEqualIID
'Fixed IID_IContextMenu/IID_IContextMenu2
'Added FreeKnownFolderDefinitionFields macro from shobjidl.h; for IKnownFolder.GetDescription

'Rev. 12
'IID additions for oleexp 4.4

'Rev. 13
'Missing IIDs ICall____

'Rev. 14
'IID additions for oleexp 4.42, 4.43

'Rev. 15
'IID additions for oleexp 4.5
'Added new FOLDERID_ values from Win10

'Rev. 16
'IID additions for oleexp 4.51

'Rev. 17
'IID additions for oleexp 4.6

'Rev. 18
'IID additions for oleexp 4.61

Public Declare Function IsEqualIID Lib "ole32" Alias "IsEqualGUID" (riid1 As UUID, riid2 As UUID) As Long


Public Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub
Public Sub DEFINE_OLEGUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer)
  DEFINE_UUID Name, L, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub
Public Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
  Name.pid = pid
End Sub

Public Function UUID_NULL() As UUID
Static bSet As Boolean
Static iid As UUID
If bSet = False Then
  With iid
    .Data1 = 0
    .Data2 = 0
    .Data3 = 0
    .Data4(0) = 0
    .Data4(1) = 0
    .Data4(2) = 0
    .Data4(3) = 0
    .Data4(4) = 0
    .Data4(5) = 0
    .Data4(6) = 0
    .Data4(7) = 0
  End With
End If
bSet = True
UUID_NULL = iid
End Function
Public Function GUIDToString(tg As UUID, Optional bBrack As Boolean = True) As String
'StringFromGUID2 never works, even "working" code from vbaccelerator AND MSDN
GUIDToString = Right$("00000000" & Hex$(tg.Data1), 8) & "-" & Right$("0000" & Hex$(tg.Data2), 4) & "-" & Right$("0000" & Hex$(tg.Data3), 4) & _
"-" & Right$("00" & Hex$(CLng(tg.Data4(0))), 2) & Right$("00" & Hex$(CLng(tg.Data4(1))), 2) & "-" & Right$("00" & Hex$(CLng(tg.Data4(2))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(3))), 2) & Right$("00" & Hex$(CLng(tg.Data4(4))), 2) & Right$("00" & Hex$(CLng(tg.Data4(5))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(6))), 2) & Right$("00" & Hex$(CLng(tg.Data4(7))), 2)
If bBrack Then GUIDToString = "{" & GUIDToString & "}"
End Function


'====================================================
'IIDs added in Rev. 8
'====================================================
Public Function IID_IShellExtInit() As UUID
'{000214E8-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214E8, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExtInit = iid
End Function
Public Function IID_IShellExecuteHookA() As UUID
'{000214F5-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214F5, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExecuteHookA = iid
End Function
Public Function IID_IShellExecuteHookW() As UUID
'{000214FB-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214FB, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellExecuteHookW = iid
End Function
Public Function IID_IEnumExtraSearch() As UUID
'{0E700BE1-9DB6-11d1-A1CE-00C04FD75D13}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE700BE1, CInt(&H9DB6), CInt(&H11D1), &HA1, &HCE, &H0, &HC0, &H4F, &HD7, &H5D, &H13)
IID_IEnumExtraSearch = iid
End Function
Public Function IID_IFolderFilterSite() As UUID
'{C0A651F5-B48B-11d2-B5ED-006097C686F6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC0A651F5, CInt(&HB48B), CInt(&H11D2), &HB5, &HED, &H0, &H60, &H97, &HC6, &H86, &HF6)
IID_IFolderFilterSite = iid
End Function
Public Function IID_IFileSystemBindData() As UUID
'{01E18D10-4D8B-11d2-855D-006008059367}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1E18D10, CInt(&H4D8B), CInt(&H11D2), &H85, &H5D, &H0, &H60, &H8, &H5, &H93, &H67)
 IID_IFileSystemBindData = iid
End Function
Public Function IID_IFileSystemBindData2() As UUID
'{3acf075f-71db-4afa-81f0-3fc4fdf2a5b8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3ACF075F, CInt(&H71DB), CInt(&H4AFA), &H81, &HF0, &H3F, &HC4, &HFD, &HF2, &HA5, &HB8)
 IID_IFileSystemBindData2 = iid
End Function
Public Function IID_IObjectWithFolderEnumMode() As UUID
'{6a9d9026-0e6e-464c-b000-42ecc07de673}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6A9D9026, CInt(&HE6E), CInt(&H464C), &HB0, &H0, &H42, &HEC, &HC0, &H7D, &HE6, &H73)
 IID_IObjectWithFolderEnumMode = iid
End Function
Public Function IID_IProfferService() As UUID
'{cb728b20-f786-11ce-92ad-00aa00a74cd0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCB728B20, CInt(&HF786), CInt(&H11CE), &H92, &HAD, &H0, &HAA, &H0, &HA7, &H4C, &HD0)
IID_IProfferService = iid
End Function
Public Function IID_IPropertyUI() As UUID
'{757a7d9f-919a-4118-99d7-dbb208c8cc66}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H757A7D9F, CInt(&H919A), CInt(&H4118), &H99, &HD7, &HDB, &HB2, &H8, &HC8, &HCC, &H66)
IID_IPropertyUI = iid
End Function
Public Function IID_ICategoryProvider() As UUID
'{9af64809-5864-4c26-a720-c1f78c086ee3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9AF64809, CInt(&H5864), CInt(&H4C26), &HA7, &H20, &HC1, &HF7, &H8C, &H8, &H6E, &HE3)
IID_ICategoryProvider = iid
End Function
Public Function IID_ICategorizer() As UUID
'{a3b14589-9174-49a8-89a3-06a1ae2b9ba7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3B14589, CInt(&H9174), CInt(&H49A8), &H89, &HA3, &H6, &HA1, &HAE, &H2B, &H9B, &HA7)
IID_ICategorizer = iid
End Function
Public Function IID_IUserEventTimerCallback() As UUID
'{e9ead8e6-2a25-410e-9b58-a9fbef1dd1a2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9EAD8E6, CInt(&H2A25), CInt(&H410E), &H9B, &H58, &HA9, &HFB, &HEF, &H1D, &HD1, &HA2)
IID_IUserEventTimerCallback = iid
End Function
Public Function IID_IUserEventTimer() As UUID
'{0F504B94-6E42-42E6-99E0-E20FAFE52AB4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF504B94, CInt(&H6E42), CInt(&H42E6), &H99, &HE0, &HE2, &HF, &HAF, &HE5, &H2A, &HB4)
IID_IUserEventTimer = iid
End Function
Public Function IID_IWebWizardExtension() As UUID
'{0e6b3f66-98d1-48c0-a222-fbde74e2fbc5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6B3F66, CInt(&H98D1), CInt(&H48C0), &HA2, &H22, &HFB, &HDE, &H74, &HE2, &HFB, &HC5)
IID_IWebWizardExtension = iid
End Function
Public Function IID_IPublishingWizard() As UUID
'{aa9198bb-ccec-472d-beed-19a4f6733f7a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAA9198BB, CInt(&HCCEC), CInt(&H472D), &HBE, &HED, &H19, &HA4, &HF6, &H73, &H3F, &H7A)
IID_IPublishingWizard = iid
End Function
Public Function IID_INetCrawler() As UUID
''{49c929ee-a1b7-4c58-b539-e63be392b6f3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49C929EE, CInt(&HA1B7), CInt(&H4C58), &HB5, &H39, &HE6, &H3B, &HE3, &H92, &HB6, &HF3)
IID_INetCrawler = iid
End Function
Public Function IID_IAsyncOperation() As UUID
'{3D8B0590-F691-11d2-8EA9-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D8B0590, CInt(&HF691), CInt(&H11D2), &H8E, &HA9, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IAsyncOperation = iid
End Function
Public Function IID_ITypeInfo2() As UUID
'{00020412-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20412, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeInfo2 = iid
End Function
Public Function IID_ITypeLib() As UUID
'{00020402-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20402, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeLib = iid
End Function
'==================================================================
'End Rev. 8 Update
'==================================================================


Public Function IID_IExplorerBrowserEvents() As UUID
'{361bbdc7-e6ee-4e13-be58-58e2240c810f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H361BBDC7, CInt(&HE6EE), CInt(&H4E13), &HBE, &H58, &H58, &HE2, &H24, &HC, &H81, &HF)
IID_IExplorerBrowserEvents = iid
End Function
Public Function IID_IExplorerBrowser() As UUID
'{dfd3b6b5-c10c-4be9-85f6-a66969f402f6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDFD3B6B5, CInt(&HC10C), CInt(&H4BE9), &H85, &HF6, &HA6, &H69, &H69, &HF4, &H2, &HF6)
IID_IExplorerBrowser = iid
End Function
Public Function IID_IExplorerPaneVisibility() As UUID
'{e07010ec-bc17-44c0-97b0-46c7c95b9edc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE07010EC, CInt(&HBC17), CInt(&H44C0), &H97, &HB0, &H46, &HC7, &HC9, &H5B, &H9E, &HDC)
IID_IExplorerPaneVisibility = iid
End Function
Public Function IID_INameSpaceTreeControl() As UUID
'{028212A3-B627-47e9-8856-C14265554E4F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H28212A3, CInt(&HB627), CInt(&H47E9), &H88, &H56, &HC1, &H42, &H65, &H55, &H4E, &H4F)
IID_INameSpaceTreeControl = iid
End Function
Public Function IID_INameSpaceTreeControl2() As UUID
'{7cc7aed8-290e-49bc-8945-c1401cc9306c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7CC7AED8, CInt(&H290E), CInt(&H49BC), &H89, &H45, &HC1, &H40, &H1C, &HC9, &H30, &H6C)
IID_INameSpaceTreeControl2 = iid
End Function
Public Function IID_INameSpaceTreeControlEvents() As UUID
'{93D77985-B3D8-4484-8318-672CDDA002CE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93D77985, CInt(&HB3D8), CInt(&H4484), &H83, &H18, &H67, &H2C, &HDD, &HA0, &H2, &HCE)
IID_INameSpaceTreeControlEvents = iid
End Function
Public Function IID_INameSpaceTreeControlDropHandler() As UUID
'{F9C665D6-C2F2-4c19-BF33-8322D7352F51}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF9C665D6, CInt(&HC2F2), CInt(&H4C19), &HBF, &H33, &H83, &H22, &HD7, &H35, &H2F, &H51)
IID_INameSpaceTreeControlDropHandler = iid
End Function
Public Function IID_INameSpaceTreeAccessible() As UUID
'{71f312de-43ed-4190-8477-e9536b82350b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71F312DE, CInt(&H43ED), CInt(&H4190), &H84, &H77, &HE9, &H53, &H6B, &H82, &H35, &HB)
IID_INameSpaceTreeAccessible = iid
End Function
Public Function IID_INameSpaceTreeControlCustomDraw() As UUID
'{2D3BA758-33EE-42d5-BB7B-5F3431D86C78}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2D3BA758, CInt(&H33EE), CInt(&H42D5), &HBB, &H7B, &H5F, &H34, &H31, &HD8, &H6C, &H78)
IID_INameSpaceTreeControlCustomDraw = iid
End Function
Public Function IID_INameSpaceTreeControlFolderCapabilities() As UUID
'{e9701183-e6b3-4ff2-8568-813615fec7be}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE9701183, CInt(&HE6B3), CInt(&H4FF2), &H85, &H68, &H81, &H36, &H15, &HFE, &HC7, &HBE)
IID_INameSpaceTreeControlFolderCapabilities = iid
End Function
Public Function IID_IShellWindows() As UUID
'{85CB6900-4D95-11CF-960C-0080C7F4EE85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85CB6900, CInt(&H4D95), CInt(&H11CF), &H96, &HC, &H0, &H80, &HC7, &HF4, &HEE, &H85)
IID_IShellWindows = iid
End Function
Public Function IID_IStreamAsync() As UUID
'{fe0b6665-e0ca-49b9-a178-2b5cb48d92a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFE0B6665, CInt(&HE0CA), CInt(&H49B9), &HA1, &H78, &H2B, &H5C, &HB4, &H8D, &H92, &HA5)
IID_IStreamAsync = iid
End Function
Public Function IID_IEnumFullIDList() As UUID
'{d0191542-7954-4908-bc06-b2360bbe45ba}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD0191542, CInt(&H7954), CInt(&H4908), &HBC, &H6, &HB2, &H36, &HB, &HBE, &H45, &HBA)
IID_IEnumFullIDList = iid
End Function
Public Function IID_IShellView3() As UUID
'{ec39fa88-f8af-41c5-8421-38bed28f4673}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC39FA88, CInt(&HF8AF), CInt(&H41C5), &H84, &H21, &H38, &HBE, &HD2, &H8F, &H46, &H73)
IID_IShellView3 = iid
End Function
Public Function IID_ICommDlgBrowser() As UUID
'{000214F1-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICommDlgBrowser = iid
End Function

Public Function IID_ICommDlgBrowser2() As UUID
'{10339516-2894-11d2-9039-00C04F8EEB3E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10339516, CInt(&H2894), CInt(&H11D2), &H90, &H39, &H0, &HC0, &H4F, &H8E, &HEB, &H3E)
IID_ICommDlgBrowser2 = iid
End Function
Public Function IID_ICommDlgBrowser3() As UUID
'{c8ad25a1-3294-41ee-8165-71174bd01c57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8AD25A1, CInt(&H3294), CInt(&H41EE), &H81, &H65, &H71, &H17, &H4B, &HD0, &H1C, &H57)
IID_ICommDlgBrowser3 = iid
End Function
Public Function IID_IColumnManager() As UUID
'{d8ec27bb-3f3b-4042-b10a-4acfd924d453}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD8EC27BB, CInt(&H3F3B), CInt(&H4042), &HB1, &HA, &H4A, &HCF, &HD9, &H24, &HD4, &H53)
IID_IColumnManager = iid
End Function
Public Function IID_ITaskbarList3() As UUID
'{ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEA1AFB91, CInt(&H9E28), CInt(&H4B86), &H90, &HE9, &H9E, &H9F, &H8A, &H5E, &HEF, &HAF)
IID_ITaskbarList3 = iid
End Function
Public Function IID_ITaskbarList4() As UUID
'{c43dc798-95d1-4bea-9030-bb99e2983a1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC43DC798, CInt(&H95D1), CInt(&H4BEA), &H90, &H30, &HBB, &H99, &HE2, &H98, &H3A, &H1A)
IID_ITaskbarList4 = iid
End Function
Public Function IID_IThumbnailProvider() As UUID
'{e357fccd-a995-4576-b01f-234630154e96}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE357FCCD, CInt(&HA995), CInt(&H4576), &HB0, &H1F, &H23, &H46, &H30, &H15, &H4E, &H96)
IID_IThumbnailProvider = iid
End Function
Public Function IID_IOperationsProgressDialog() As UUID
'{0C9FB851-E5C9-43EB-A370-F0677B13874C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9FB851, CInt(&HE5C9), CInt(&H43EB), &HA3, &H70, &HF0, &H67, &H7B, &H13, &H87, &H4C)
IID_IOperationsProgressDialog = iid
End Function
Public Function IID_IFileOperationProgressSink() As UUID
'{04b0f1a7-9490-44bc-96e1-4296a31252e2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4B0F1A7, CInt(&H9490), CInt(&H44BC), &H96, &HE1, &H42, &H96, &HA3, &H12, &H52, &HE2)
IID_IFileOperationProgressSink = iid
End Function
Public Function IID_IFileOperation() As UUID
'{947aab5f-0a5c-4c13-b4d6-4bf7836fc9f8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H947AAB5F, CInt(&HA5C), CInt(&H4C13), &HB4, &HD6, &H4B, &HF7, &H83, &H6F, &HC9, &HF8)
IID_IFileOperation = iid
End Function
Public Function IID_IObjectCollection() As UUID
'{5632b1a4-e38a-400a-928a-d4cd63230295}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5632B1A4, CInt(&HE38A), CInt(&H400A), &H92, &H8A, &HD4, &HCD, &H63, &H23, &H2, &H95)
IID_IObjectCollection = iid
End Function
Public Function IID_IApplicationDestinations() As UUID
'{12337d35-94c6-48a0-bce7-6a9c69d4d600}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12337D35, CInt(&H94C6), CInt(&H48A0), &HBC, &HE7, &H6A, &H9C, &H69, &HD4, &HD6, &H0)
IID_IApplicationDestinations = iid
End Function
Public Function IID_ICustomDestinationList() As UUID
'{6332debf-87b5-4670-90c0-5e57b408a49e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6332DEBF, CInt(&H87B5), CInt(&H4670), &H90, &HC0, &H5E, &H57, &HB4, &H8, &HA4, &H9E)
IID_ICustomDestinationList = iid
End Function
Public Function IID_IModalWindow() As UUID
'{b4db1657-70d7-485e-8e3e-6fcb5a5c1802}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4DB1657, CInt(&H70D7), CInt(&H485E), &H8E, &H3E, &H6F, &HCB, &H5A, &H5C, &H18, &H2)
IID_IModalWindow = iid
End Function
Public Function IID_IFileDialogEvents() As UUID
'{973510db-7d7f-452b-8975-74a85828d354}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H973510DB, CInt(&H7D7F), CInt(&H452B), &H89, &H75, &H74, &HA8, &H58, &H28, &HD3, &H54)
IID_IFileDialogEvents = iid
End Function
Public Function IID_IShellItemFilter() As UUID
'{2659B475-EEB8-48b7-8F07-B378810F48CF}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2659B475, CInt(&HEEB8), CInt(&H48B7), &H8F, &H7, &HB3, &H78, &H81, &HF, &H48, &HCF)
IID_IShellItemFilter = iid
End Function
Public Function IID_IFileDialog() As UUID
'{42f85136-db7e-439c-85f1-e4075d135fc8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H42F85136, CInt(&HDB7E), CInt(&H439C), &H85, &HF1, &HE4, &H7, &H5D, &H13, &H5F, &HC8)
IID_IFileDialog = iid
End Function
Public Function IID_IFileSaveDialog() As UUID
'{84bccd23-5fde-4cdb-aea4-af64b83d78ab}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84BCCD23, CInt(&H5FDE), CInt(&H4CDB), &HAE, &HA4, &HAF, &H64, &HB8, &H3D, &H78, &HAB)
IID_IFileSaveDialog = iid
End Function
Public Function IID_IFileOpenDialog() As UUID
'{d57c7288-d4ad-4768-be02-9d969532d960}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD57C7288, CInt(&HD4AD), CInt(&H4768), &HBE, &H2, &H9D, &H96, &H95, &H32, &HD9, &H60)
IID_IFileOpenDialog = iid
End Function
Public Function IID_IFileDialogControlEvents() As UUID
'{36116642-D713-4b97-9B83-7484A9D00433}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H36116642, CInt(&HD713), CInt(&H4B97), &H9B, &H83, &H74, &H84, &HA9, &HD0, &H4, &H33)
IID_IFileDialogControlEvents = iid
End Function
Public Function IID_IFileDialog2() As UUID
'{61744fc7-85b5-4791-a9b0-272276309b13}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H61744FC7, CInt(&H85B5), CInt(&H4791), &HA9, &HB0, &H27, &H22, &H76, &H30, &H9B, &H13)
IID_IFileDialog2 = iid
End Function
Public Function IID_IShellMenuCallback() As UUID
'{4CA300A1-9B8D-11d1-8B22-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CA300A1, CInt(&H9B8D), CInt(&H11D1), &H8B, &H22, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
IID_IShellMenuCallback = iid
End Function
Public Function IID_IAssocHandlerInvoker() As UUID
'{92218CAB-ECAA-4335-8133-807FD234C2EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92218CAB, CInt(&HECAA), CInt(&H4335), &H81, &H33, &H80, &H7F, &HD2, &H34, &HC2, &HEE)
IID_IAssocHandlerInvoker = iid
End Function
Public Function IID_IAssocHandler() As UUID
'{F04061AC-1659-4a3f-A954-775AA57FC083}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF04061AC, CInt(&H1659), CInt(&H4A3F), &HA9, &H54, &H77, &H5A, &HA5, &H7F, &HC0, &H83)
IID_IAssocHandler = iid
End Function
Public Function IID_IEnumAssocHandlers() As UUID
'{973810ae-9599-4b88-9e4d-6ee98c9552da}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H973810AE, CInt(&H9599), CInt(&H4B88), &H9E, &H4D, &H6E, &HE9, &H8C, &H95, &H52, &HDA)
IID_IEnumAssocHandlers = iid
End Function
Public Function IID_INamespaceWalkCB() As UUID
'{d92995f8-cf5e-4a76-bf59-ead39ea2b97e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD92995F8, CInt(&HCF5E), CInt(&H4A76), &HBF, &H59, &HEA, &HD3, &H9E, &HA2, &HB9, &H7E)
IID_INamespaceWalkCB = iid
End Function
Public Function IID_INamespaceWalkCB2() As UUID
'{7ac7492b-c38e-438a-87db-68737844ff70}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7AC7492B, CInt(&HC38E), CInt(&H438A), &H87, &HDB, &H68, &H73, &H78, &H44, &HFF, &H70)
IID_INamespaceWalkCB2 = iid
End Function
Public Function IID_INamespaceWalk() As UUID
'{57ced8a7-3f4a-432c-9350-30f24483f74f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57CED8A7, CInt(&H3F4A), CInt(&H432C), &H93, &H50, &H30, &HF2, &H44, &H83, &HF7, &H4F)
IID_INamespaceWalk = iid
End Function
Public Function IID_IUserNotificationCallback() As UUID
'{19108294-0441-4AFF-8013-FA0A730B0BEA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H19108294, CInt(&H441), CInt(&H4AFF), &H80, &H13, &HFA, &HA, &H73, &HB, &HB, &HEA)
IID_IUserNotificationCallback = iid
End Function
Public Function IID_IUserNotification2() As UUID
'{215913CC-57EB-4FAB-AB5A-E5FA7BEA2A6C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H215913CC, CInt(&H57EB), CInt(&H4FAB), &HAB, &H5A, &HE5, &HFA, &H7B, &HEA, &H2A, &H6C)
IID_IUserNotification2 = iid
End Function
Public Function IID_ITransferAdviseSink() As UUID
'{d594d0d8-8da7-457b-b3b4-ce5dbaac0b88}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD594D0D8, CInt(&H8DA7), CInt(&H457B), &HB3, &HB4, &HCE, &H5D, &HBA, &HAC, &HB, &H88)
IID_ITransferAdviseSink = iid
End Function
Public Function IID_IObjectWithPropertyKey() As UUID
'{fc0ca0a7-c316-4fd2-9031-3e628e6d4f23}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC0CA0A7, CInt(&HC316), CInt(&H4FD2), &H90, &H31, &H3E, &H62, &H8E, &H6D, &H4F, &H23)
IID_IObjectWithPropertyKey = iid
End Function
Public Function IID_IPropertyChange() As UUID
'{f917bc8a-1bba-4478-a245-1bde03eb9431}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF917BC8A, CInt(&H1BBA), CInt(&H4478), &HA2, &H45, &H1B, &HDE, &H3, &HEB, &H94, &H31)
IID_IPropertyChange = iid
End Function
Public Function IID_IPropertyChangeArray() As UUID
'{380f5cad-1b5e-42f2-805d-637fd392d31e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H380F5CAD, CInt(&H1B5E), CInt(&H42F2), &H80, &H5D, &H63, &H7F, &HD3, &H92, &HD3, &H1E)
IID_IPropertyChangeArray = iid
End Function
Public Function IID_IPropertyDescription2() As UUID
'{57d2eded-5062-400e-b107-5dae79fe57a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H57D2EDED, CInt(&H5062), CInt(&H400E), &HB1, &H7, &H5D, &HAE, &H79, &HFE, &H57, &HA6)
IID_IPropertyDescription2 = iid
End Function
Public Function IID_IPropertyDescriptionSearchInfo() As UUID
'{078f91bd-29a2-440f-924e-46a291524520}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H78F91BD, CInt(&H29A2), CInt(&H440F), &H92, &H4E, &H46, &HA2, &H91, &H52, &H45, &H20)
IID_IPropertyDescriptionSearchInfo = iid
End Function
Public Function IID_IPropertyDescriptionRelatedPropertyInfo() As UUID
'{507393f4-2a3d-4a60-b59e-d9c75716c2dd}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H507393F4, CInt(&H2A3D), CInt(&H4A60), &HB5, &H9E, &HD9, &HC7, &H57, &H16, &HC2, &HDD)
IID_IPropertyDescriptionRelatedPropertyInfo = iid
End Function
Public Function IID_IPropertyEnumType() As UUID
'{11e1fbf9-2d56-4a6b-8db3-7cd193a471f2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11E1FBF9, CInt(&H2D56), CInt(&H4A6B), &H8D, &HB3, &H7C, &HD1, &H93, &HA4, &H71, &HF2)
IID_IPropertyEnumType = iid
End Function
Public Function IID_IPropertyEnumType2() As UUID
'{9b6e051c-5ddd-4321-9070-fe2acb55e794}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B6E051C, CInt(&H5DDD), CInt(&H4321), &H90, &H70, &HFE, &H2A, &HCB, &H55, &HE7, &H94)
IID_IPropertyEnumType2 = iid
End Function
Public Function IID_IPropertyEnumTypeList() As UUID
'{a99400f4-3d84-4557-94ba-1242fb2cc9a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA99400F4, CInt(&H3D84), CInt(&H4557), &H94, &HBA, &H12, &H42, &HFB, &H2C, &HC9, &HA6)
IID_IPropertyEnumTypeList = iid
End Function
Public Function IID_IPropertyStoreFactory() As UUID
'{bc110b6d-57e8-4148-a9c6-91015ab2f3a5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC110B6D, CInt(&H57E8), CInt(&H4148), &HA9, &HC6, &H91, &H1, &H5A, &HB2, &HF3, &HA5)
IID_IPropertyStoreFactory = iid
End Function
Public Function IID_IPropertyStoreCapabilities() As UUID
'{c8e2d566-186e-4d49-bf41-6909ead56acc}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC8E2D566, CInt(&H186E), CInt(&H4D49), &HBF, &H41, &H69, &H9, &HEA, &HD5, &H6A, &HCC)
IID_IPropertyStoreCapabilities = iid
End Function
Public Function IID_IPropertyStoreCache() As UUID
'{3017056d-9a91-4e90-937d-746c72abbf4f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3017056D, CInt(&H9A91), CInt(&H4E90), &H93, &H7D, &H74, &H6C, &H72, &HAB, &HBF, &H4F)
IID_IPropertyStoreCache = iid
End Function
Public Function IID_INamedPropertyStore() As UUID
'{71604b0f-97b0-4764-8577-2f13e98a1422}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H71604B0F, CInt(&H97B0), CInt(&H4764), &H85, &H77, &H2F, &H13, &HE9, &H8A, &H14, &H22)
 IID_INamedPropertyStore = iid
End Function
Public Function IID_IPropertyDescriptionAliasInfo() As UUID
'{f67104fc-2af9-46fd-b32d-243c1404f3d1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF67104FC, CInt(&H2AF9), CInt(&H46FD), &HB3, &H2D, &H24, &H3C, &H14, &H4, &HF3, &HD1)
 IID_IPropertyDescriptionAliasInfo = iid
End Function
Public Function IID_IAutoComplete() As UUID
'{00bb2762-6a77-11d0-a535-00c04fd7d062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB2762, CInt(&H6A77), CInt(&H11D0), &HA5, &H35, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IAutoComplete = iid
End Function
Public Function IID_IAutoComplete2() As UUID
'{EAC04BC0-3791-11d2-BB95-0060977B464C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAC04BC0, CInt(&H3791), CInt(&H11D2), &HBB, &H95, &H0, &H60, &H97, &H7B, &H46, &H4C)
IID_IAutoComplete2 = iid
End Function
Public Function IID_IEnumACString() As UUID
'{8E74C210-CF9D-4eaf-A403-7356428F0A5A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8E74C210, CInt(&HCF9D), CInt(&H4EAF), &HA4, &H3, &H73, &H56, &H42, &H8F, &HA, &H5A)
IID_IEnumACString = iid
End Function
Public Function IID_IACList() As UUID
'{77A130B0-94FD-11D0-A544-00C04FD7d062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77A130B0, CInt(&H94FD), CInt(&H11D0), &HA5, &H44, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IACList = iid
End Function
Public Function IID_IACList2() As UUID
'{470141a0-5186-11d2-bbb6-0060977b464c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H470141A0, CInt(&H5186), CInt(&H11D2), &HBB, &HB6, &H0, &H60, &H97, &H7B, &H46, &H4C)
IID_IACList2 = iid
End Function
Public Function IID_IBindCtx() As UUID
'{0000000e-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IBindCtx = iid
End Function
Public Function IID_IRunningObjectTable() As UUID
'{00000010-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRunningObjectTable = iid
End Function
Public Function IID_ICatRegister() As UUID
'{0002E012-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E012, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICatRegister = iid
End Function
Public Function IID_ICatInformation() As UUID
'{0002E013-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E013, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICatInformation = iid
End Function
Public Function IID_ICreateTypeInfo() As UUID
'{00020405-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20405, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeInfo = iid
End Function
Public Function IID_ICreateTypeInfo2() As UUID
'{0002040E-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2040E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeInfo2 = iid
End Function
Public Function IID_ICreateTypeLib() As UUID
'{00020406-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20406, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeLib = iid
End Function
Public Function IID_ICreateTypeLib2() As UUID
'{0002040F-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2040F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICreateTypeLib2 = iid
End Function
Public Function IID_IDocHostUIHandler() As UUID
'{bd3f23c0-d43e-11cf-893b-00aa00bdce1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD3F23C0, CInt(&HD43E), CInt(&H11CF), &H89, &H3B, &H0, &HAA, &H0, &HBD, &HCE, &H1A)
IID_IDocHostUIHandler = iid
End Function
Public Function IID_IDocHostUIHandler2() As UUID
'{3050f6d0-98b5-11cf-bb82-00aa00bdce0b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3050F6D0, CInt(&H98B5), CInt(&H11CF), &HBB, &H82, &H0, &HAA, &H0, &HBD, &HCE, &HB)
IID_IDocHostUIHandler2 = iid
End Function
Public Function IID_ICustomDoc() As UUID
'{3050f3f0-98b5-11cf-bb82-00aa00bdce0b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3050F3F0, CInt(&H98B5), CInt(&H11CF), &HBB, &H82, &H0, &HAA, &H0, &HBD, &HCE, &HB)
IID_ICustomDoc = iid
End Function
Public Function IID_IDocHostShowUI() As UUID
'{c4d244b0-d43e-11cf-893b-00aa00bdce1a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC4D244B0, CInt(&HD43E), CInt(&H11CF), &H89, &H3B, &H0, &HAA, &H0, &HBD, &HCE, &H1A)
IID_IDocHostShowUI = iid
End Function
Public Function IID_IAdviseSink() As UUID
'{0000010f-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IAdviseSink = iid
End Function
Public Function IID_IInputObject() As UUID
'{68284faa-6a48-11d0-8c78-00c04fd918b4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H68284FAA, CInt(&H6A48), CInt(&H11D0), &H8C, &H78, &H0, &HC0, &H4F, &HD9, &H18, &HB4)
IID_IInputObject = iid
End Function
Public Function IID_IDeskBand() As UUID
'{EB0FE172-1A3A-11D0-89B3-00A0C90A90AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEB0FE172, CInt(&H1A3A), CInt(&H11D0), &H89, &HB3, &H0, &HA0, &HC9, &HA, &H90, &HAC)
IID_IDeskBand = iid
End Function
Public Function IID_IDockingWindow() As UUID
'{012dd920-7b26-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12DD920, CInt(&H7B26), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindow = iid
End Function
Public Function IID_IDockingWindowSite() As UUID
'{2a342fc2-7b26-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A342FC2, CInt(&H7B26), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindowSite = iid
End Function
Public Function IID_IDockingWindowFrame() As UUID
'{47d2657a-7b27-11d0-8ca9-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H47D2657A, CInt(&H7B27), CInt(&H11D0), &H8C, &HA9, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IDockingWindowFrame = iid
End Function
Public Function IID_IInputObjectSite() As UUID
'{f1db8392-7331-11d0-8c99-00a0c92dbfe8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF1DB8392, CInt(&H7331), CInt(&H11D0), &H8C, &H99, &H0, &HA0, &HC9, &H2D, &HBF, &HE8)
IID_IInputObjectSite = iid
End Function
Public Function IID_IEnumSTATPROPSTG() As UUID
'{00000139-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H139, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATPROPSTG = iid
End Function
Public Function IID_IEnumSTATPROPSETSTG() As UUID
'{0000013B-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H13B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATPROPSETSTG = iid
End Function
Public Function IID_IEnumSTATSTG() As UUID
'{0000000d-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATSTG = iid
End Function
Public Function IID_IEnumSTATDATA() As UUID
'{00000105-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H105, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumSTATDATA = iid
End Function
Public Function IID_IEnumString() As UUID
'{00000101-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H101, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumString = iid
End Function
Public Function IID_IEnumMoniker() As UUID
'{00000102-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H102, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumMoniker = iid
End Function
Public Function IID_IEnumFORMATETC() As UUID
'{00000103-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H103, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumFORMATETC = iid
End Function
Public Function IID_IEnumUnknown() As UUID
'{00000100-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H100, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumUnknown = iid
End Function
Public Function IID_IEnumOLEVERB() As UUID
'{00000104-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H104, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumOLEVERB = iid
End Function
Public Function IID_IEnumGUID() As UUID
'{0002E000-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E000, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumGUID = iid
End Function
Public Function IID_IEnumCATEGORYINFO() As UUID
'{0002E011-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E011, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumCATEGORYINFO = iid
End Function
Public Function IID_IEnumVARIANT() As UUID
'{00020404-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20404, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumVARIANT = iid
End Function
Public Function IID_IEnumConnections() As UUID
'{B196B287-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B287, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IEnumConnections = iid
End Function
Public Function IID_IEnumConnectionPoints() As UUID
'{B196B285-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B285, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IEnumConnectionPoints = iid
End Function
Public Function IID_IErrorInfo() As UUID
'{1CF2B120-547D-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1CF2B120, CInt(&H547D), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_IErrorInfo = iid
End Function
Public Function IID_ICreateErrorInfo() As UUID
'{22F03340-547D-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F03340, CInt(&H547D), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_ICreateErrorInfo = iid
End Function
Public Function IID_ISupportErrorInfo() As UUID
'{DF0B3D60-548F-101B-8E65-08002B2BD119}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDF0B3D60, CInt(&H548F), CInt(&H101B), &H8E, &H65, &H8, &H0, &H2B, &H2B, &HD1, &H19)
IID_ISupportErrorInfo = iid
End Function
Public Function IID_IEmptyVolumeCacheCallBack() As UUID
'{6E793361-73C6-11D0-8469-00AA00442901}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E793361, CInt(&H73C6), CInt(&H11D0), &H84, &H69, &H0, &HAA, &H0, &H44, &H29, &H1)
IID_IEmptyVolumeCacheCallBack = iid
End Function
Public Function IID_IEmptyVolumeCache() As UUID
'{8FCE5227-04DA-11d1-A004-00805F8ABE06}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FCE5227, CInt(&H4DA), CInt(&H11D1), &HA0, &H4, &H0, &H80, &H5F, &H8A, &HBE, &H6)
IID_IEmptyVolumeCache = iid
End Function
Public Function IID_IEmptyVolumeCache2() As UUID
'{02b7e3ba-4db3-11d2-b2d9-00c04f8eec8c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2B7E3BA, CInt(&H4DB3), CInt(&H11D2), &HB2, &HD9, &H0, &HC0, &H4F, &H8E, &HEC, &H8C)
IID_IEmptyVolumeCache2 = iid
End Function
Public Function IID_IPublishedApp() As UUID
'{1BC752E0-9046-11D1-B8B3-006008059382}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1BC752E0, CInt(&H9046), CInt(&H11D1), &HB8, &HB3, &H0, &H60, &H8, &H5, &H93, &H82)
IID_IPublishedApp = iid
End Function
Public Function IID_IPublishedApp2() As UUID
'{12B81347-1B3A-4A04-AA61-3F768B67FD7E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12B81347, CInt(&H1B3A), CInt(&H4A04), &HAA, &H61, &H3F, &H76, &H8B, &H67, &HFD, &H7E)
IID_IPublishedApp2 = iid
End Function
Public Function IID_IEnumPublishedApps() As UUID
'{0B124F8C-91F0-11D1-B8B5-006008059382}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB124F8C, CInt(&H91F0), CInt(&H11D1), &HB8, &HB5, &H0, &H60, &H8, &H5, &H93, &H82)
IID_IEnumPublishedApps = iid
End Function
Public Function IID_IShellBrowser() As UUID
'{000214E2-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellBrowser = iid
End Function
Public Function IID_IProgressDialog() As UUID
'{EBBC7C04-315E-11d2-B62F-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEBBC7C04, CInt(&H315E), CInt(&H11D2), &HB6, &H2F, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IProgressDialog = iid
End Function
Public Function IID_IMoniker() As UUID
'{0000000f-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMoniker = iid
End Function
Public Function IID_IHlink() As UUID
'{79eac9c3-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C3, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlink = iid
End Function
Public Function IID_IHlinkSite() As UUID
'{79eac9c2-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C2, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkSite = iid
End Function
Public Function IID_IHlinkTarget() As UUID
'{79eac9c4-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C4, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkTarget = iid
End Function
Public Function IID_IHlinkFrame() As UUID
'{79eac9c5-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C5, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkFrame = iid
End Function
Public Function IID_IEnumHLITEM() As UUID
'{79eac9c6-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C6, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IEnumHLITEM = iid
End Function
Public Function IID_IHlinkBrowseContext() As UUID
'{79eac9c7-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C7, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHlinkBrowseContext = iid
End Function
Public Function IID_IDiscRecorder() As UUID
'{85AC9776-CA88-4cf2-894E-09598C078A41}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85AC9776, CInt(&HCA88), CInt(&H4CF2), &H89, &H4E, &H9, &H59, &H8C, &H7, &H8A, &H41)
IID_IDiscRecorder = iid
End Function
Public Function IID_IEnumDiscRecorders() As UUID
'{9B1921E1-54AC-11d3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B1921E1, CInt(&H54AC), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IEnumDiscRecorders = iid
End Function
Public Function IID_IEnumDiscMasterFormats() As UUID
'{DDF445E1-54BA-11d3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDF445E1, CInt(&H54BA), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IEnumDiscMasterFormats = iid
End Function
Public Function IID_IRedbookDiscMaster() As UUID
'{E3BC42CD-4E5C-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3BC42CD, CInt(&H4E5C), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IRedbookDiscMaster = iid
End Function
Public Function IID_IJolietDiscMaster() As UUID
'{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE3BC42CE, CInt(&H4E5C), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IJolietDiscMaster = iid
End Function
Public Function IID_IDiscMasterProgressEvents() As UUID
'{EC9E51C1-4E5D-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEC9E51C1, CInt(&H4E5D), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IDiscMasterProgressEvents = iid
End Function
Public Function IID_IDiscMaster() As UUID
'{520CCA62-51A5-11D3-9144-00104BA11C5E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H520CCA62, CInt(&H51A5), CInt(&H11D3), &H91, &H44, &H0, &H10, &H4B, &HA1, &H1C, &H5E)
IID_IDiscMaster = iid
End Function
Public Function IID_IOleInPlaceUIWindow() As UUID
'{00000115-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H115, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceUIWindow = iid
End Function
Public Function IID_IOleInPlaceActiveObject() As UUID
'{00000117-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H117, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceActiveObject = iid
End Function
Public Function IID_IOleInPlaceSite() As UUID
'{00000119-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H119, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceSite = iid
End Function
Public Function IID_IOleInPlaceFrame() As UUID
'{00000116-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H116, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceFrame = iid
End Function
Public Function IID_IOleInPlaceObject() As UUID
'{00000113-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H113, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleInPlaceObject = iid
End Function
Public Function IID_IOleControlSite() As UUID
'{B196B289-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B289, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IOleControlSite = iid
End Function
Public Function IID_ILockBytes() As UUID
'{0000000a-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ILockBytes = iid
End Function
Public Function IID_IFillLockBytes() As UUID
'{99caf010-415e-11cf-8814-00aa00b569f5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H99CAF010, CInt(&H415E), CInt(&H11CF), &H88, &H14, &H0, &HAA, &H0, &HB5, &H69, &HF5)
IID_IFillLockBytes = iid
End Function
Public Function IID_IMalloc() As UUID
'{00000002-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMalloc = iid
End Function
Public Function IID_IMarshal() As UUID
'{00000003-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IMarshal = iid
End Function
Public Function IID_IObjectSafety() As UUID
'{CB5BDC81-93C1-11cf-8F20-00805F2CD064}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCB5BDC81, CInt(&H93C1), CInt(&H11CF), &H8F, &H20, &H0, &H80, &H5F, &H2C, &HD0, &H64)
IID_IObjectSafety = iid
End Function
Public Function IID_IOleDocument() As UUID
'{b722bcc5-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC5, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocument = iid
End Function
Public Function IID_IOleDocumentSite() As UUID
'{b722bcc7-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC7, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocumentSite = iid
End Function
Public Function IID_IOleDocumentView() As UUID
'{b722bcc6-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC6, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleDocumentView = iid
End Function
Public Function IID_IEnumOleDocumentViews() As UUID
'{b722bcc8-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC8, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IEnumOleDocumentViews = iid
End Function
Public Function IID_IContinueCallback() As UUID
'{b722bcca-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCCA, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IContinueCallback = iid
End Function
Public Function IID_IPrint() As UUID
'{b722bcc9-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCC9, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IPrint = iid
End Function
Public Function IID_IOleClientSite() As UUID
'{00000118-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H118, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleClientSite = iid
End Function
Public Function IID_IParseDisplayName() As UUID
'{0000011A-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IParseDisplayName = iid
End Function
Public Function IID_IOleContainer() As UUID
'{0000011B-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleContainer = iid
End Function
Public Function IID_IOleObject() As UUID
'{00000112-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H112, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleObject = iid
End Function
Public Function IID_IOleCache() As UUID
'{0000011e-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H11E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleCache = iid
End Function
Public Function IID_IOleControl() As UUID
'{B196B288-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B288, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IOleControl = iid
End Function
Public Function IID_IOleCommandTarget() As UUID
'{b722bccb-4e68-101b-a2bc-00aa00404770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB722BCCB, CInt(&H4E68), CInt(&H101B), &HA2, &HBC, &H0, &HAA, &H0, &H40, &H47, &H70)
IID_IOleCommandTarget = iid
End Function
Public Function IID_IServiceProvider() As UUID
'{6d5140c1-7436-11ce-8034-00aa006009fa}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D5140C1, CInt(&H7436), CInt(&H11CE), &H80, &H34, &H0, &HAA, &H0, &H60, &H9, &HFA)
IID_IServiceProvider = iid
End Function
Public Function IID_ISpecifyPropertyPages() As UUID
'{B196B28B-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B28B, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_ISpecifyPropertyPages = iid
End Function
Public Function IID_IOleWindow() As UUID
'{00000114-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H114, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IOleWindow = iid
End Function
Public Function IID_IObjectWithSite() As UUID
'{FC4801A3-2BA9-11CF-A229-00AA003D7352}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC4801A3, CInt(&H2BA9), CInt(&H11CF), &HA2, &H29, &H0, &HAA, &H0, &H3D, &H73, &H52)
IID_IObjectWithSite = iid
End Function
Public Function IID_IPersist() As UUID
'{0000010c-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10C, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersist = iid
End Function
Public Function IID_IPersistStream() As UUID
'{00000109-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H109, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistStream = iid
End Function
Public Function IID_IPersistStreamInit() As UUID
'{7FD52380-4E07-101B-AE2D-08002B2EC713}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FD52380, CInt(&H4E07), CInt(&H101B), &HAE, &H2D, &H8, &H0, &H2B, &H2E, &HC7, &H13)
IID_IPersistStreamInit = iid
End Function
Public Function IID_IPersistFile() As UUID
'{0000010b-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10B, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistFile = iid
End Function
Public Function IID_IPersistStorage() As UUID
'{0000010a-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistStorage = iid
End Function
Public Function IID_IPersistPropertyBag() As UUID
'{37D84F60-42CB-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H37D84F60, CInt(&H42CB), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
IID_IPersistPropertyBag = iid
End Function
Public Function IID_IPersistPropertyBag2() As UUID
'{22F55881-280B-11d0-A8A9-00A0C90C2004}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F55881, CInt(&H280B), CInt(&H11D0), &HA8, &HA9, &H0, &HA0, &HC9, &HC, &H20, &H4)
IID_IPersistPropertyBag2 = iid
End Function
Public Function IID_IPersistMemory() As UUID
'{BD1AE5E0-A6AE-11CE-BD37-504200C10000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBD1AE5E0, CInt(&HA6AE), CInt(&H11CE), &HBD, &H37, &H50, &H42, &H0, &HC1, &H0, &H0)
IID_IPersistMemory = iid
End Function
Public Function IID_IPersistMoniker() As UUID
'{79eac9c9-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C9, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IPersistMoniker = iid
End Function
Public Function IID_IPerPropertyBrowsing() As UUID
'{376BD3AA-3845-101B-84ED-08002B2EC713}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H376BD3AA, CInt(&H3845), CInt(&H101B), &H84, &HED, &H8, &H0, &H2B, &H2E, &HC7, &H13)
IID_IPerPropertyBrowsing = iid
End Function
Public Function IID_IErrorLog() As UUID
'{3127CA40-446E-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3127CA40, CInt(&H446E), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
IID_IErrorLog = iid
End Function
Public Function IID_IPropertyBag2() As UUID
'{22F55882-280B-11d0-A8A9-00A0C90C2004}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22F55882, CInt(&H280B), CInt(&H11D0), &HA8, &HA9, &H0, &HA0, &HC9, &HC, &H20, &H4)
IID_IPropertyBag2 = iid
End Function
Public Function IID_IPropertyNotifySink() As UUID
'{9BFBBC02-EFF1-101A-84ED-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9BFBBC02, CInt(&HEFF1), CInt(&H101A), &H84, &HED, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IPropertyNotifySink = iid
End Function
Public Function IID_IRecordInfo() As UUID
'{0000002F-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRecordInfo = iid
End Function
Public Function IID_IRichEditOle() As UUID
'{00020D00-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20D00, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRichEditOle = iid
End Function
Public Function IID_IRichEditOleCallback() As UUID
'{00020D03-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20D03, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IRichEditOleCallback = iid
End Function
Public Function IID_IInternetSecurityMgrSite() As UUID
'{79eac9ed-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9ED, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSecurityMgrSite = iid
End Function
Public Function IID_IInternetSecurityManager() As UUID
'{79eac9ee-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EE, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSecurityManager = iid
End Function
Public Function IID_IInternetHostSecurityManager() As UUID
'{3af280b6-cb3f-11d0-891e-00c04fb6bfc4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AF280B6, CInt(&HCB3F), CInt(&H11D0), &H89, &H1E, &H0, &HC0, &H4F, &HB6, &HBF, &HC4)
IID_IInternetHostSecurityManager = iid
End Function
Public Function IID_IInternetZoneManager() As UUID
'{79eac9ef-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EF, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetZoneManager = iid
End Function
Public Function IID_IPersistFolder() As UUID
'{000214EA-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IPersistFolder = iid
End Function
Public Function IID_IPersistFolder2() As UUID
'{1AC3D9F0-175C-11d1-95BE-00609797EA4F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1AC3D9F0, CInt(&H175C), CInt(&H11D1), &H95, &HBE, &H0, &H60, &H97, &H97, &HEA, &H4F)
IID_IPersistFolder2 = iid
End Function
Public Function IID_IPersistFolder3() As UUID
'{CEF04FDF-FE72-11d2-87a5-00c04f6837cf}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCEF04FDF, CInt(&HFE72), CInt(&H11D2), &H87, &HA5, &H0, &HC0, &H4F, &H68, &H37, &HCF)
 IID_IPersistFolder3 = iid
End Function
Public Function IID_IPersistIDList() As UUID
'{1079acfc-29bd-11d3-8e0d-00c04f6837d5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1079ACFC, CInt(&H29BD), CInt(&H11D3), &H8E, &HD, &H0, &HC0, &H4F, &H68, &H37, &HD5)
IID_IPersistIDList = iid
End Function
Public Function IID_IShellView2() As UUID
'{88E39E80-3578-11CF-AE69-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88E39E80, CInt(&H3578), CInt(&H11CF), &HAE, &H69, &H8, &H0, &H2B, &H2E, &H12, &H62)
IID_IShellView2 = iid
End Function
Public Function IID_IEnumIDList() As UUID
'{000214F2-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IEnumIDList = iid
End Function
Public Function IID_IShellIcon() As UUID
'{000214E5-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E5, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellIcon = iid
End Function
Public Function IID_IShellLinkA() As UUID
'{000214EE-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EE, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellLinkA = iid
End Function
Public Function IID_IShellLinkW() As UUID
'{000214F9-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F9, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellLinkW = iid
End Function
Public Function IID_IActionProgressDialog() As UUID
'{49ff1172-eadc-446d-9285-156453a6431c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49FF1172, CInt(&HEADC), CInt(&H446D), &H92, &H85, &H15, &H64, &H53, &HA6, &H43, &H1C)
IID_IActionProgressDialog = iid
End Function
Public Function IID_IHWEventHandler() As UUID
'{C1FB73D0-EC3A-4ba2-B512-8CDB9187B6D1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC1FB73D0, CInt(&HEC3A), CInt(&H4BA2), &HB5, &H12, &H8C, &HDB, &H91, &H87, &HB6, &HD1)
IID_IHWEventHandler = iid
End Function
Public Function IID_IQueryCancelAutoPlay() As UUID
'{DDEFE873-6997-4e68-BE26-39B633ADBE12}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDEFE873, CInt(&H6997), CInt(&H4E68), &HBE, &H26, &H39, &HB6, &H33, &HAD, &HBE, &H12)
IID_IQueryCancelAutoPlay = iid
End Function
Public Function IID_IActionProgress() As UUID
'{49ff1173-eadc-446d-9285-156453a6431c}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H49FF1173, CInt(&HEADC), CInt(&H446D), &H92, &H85, &H15, &H64, &H53, &HA6, &H43, &H1C)
IID_IActionProgress = iid
End Function
Public Function IID_IQueryContinue() As UUID
'{7307055c-b24a-486b-9f25-163e597a28a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7307055C, CInt(&HB24A), CInt(&H486B), &H9F, &H25, &H16, &H3E, &H59, &H7A, &H28, &HA9)
IID_IQueryContinue = iid
End Function
Public Function IID_IUserNotification() As UUID
'{ba9711ba-5893-4787-a7e1-41277151550b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBA9711BA, CInt(&H5893), CInt(&H4787), &HA7, &HE1, &H41, &H27, &H71, &H51, &H55, &HB)
IID_IUserNotification = iid
End Function
Public Function IID_ITaskbarList() As UUID
'{56FDF342-FD6D-11d0-958A-006097C9A090}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56FDF342, CInt(&HFD6D), CInt(&H11D0), &H95, &H8A, &H0, &H60, &H97, &HC9, &HA0, &H90)
IID_ITaskbarList = iid
End Function
Public Function IID_ITaskbarList2() As UUID
'{602D4995-B13A-429b-A66E-1935E44F4317}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H602D4995, CInt(&HB13A), CInt(&H429B), &HA6, &H6E, &H19, &H35, &HE4, &H4F, &H43, &H17)
IID_ITaskbarList2 = iid
End Function
Public Function IID_IActiveDesktop() As UUID
'{F490EB00-1240-11D1-9888-006097DEACF9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF490EB00, CInt(&H1240), CInt(&H11D1), &H98, &H88, &H0, &H60, &H97, &HDE, &HAC, &HF9)
IID_IActiveDesktop = iid
End Function
Public Function IID_ICDBurn() As UUID
'{3d73a659-e5d0-4d42-afc0-5121ba425c8d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D73A659, CInt(&HE5D0), CInt(&H4D42), &HAF, &HC0, &H51, &H21, &HBA, &H42, &H5C, &H8D)
IID_ICDBurn = iid
End Function
Public Function IID_ICDBurnExt() As UUID
'{2271dcca-74fc-4414-8fb7-c56b05ace2d7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2271DCCA, CInt(&H74FC), CInt(&H4414), &H8F, &HB7, &HC5, &H6B, &H5, &HAC, &HE2, &HD7)
 IID_ICDBurnExt = iid
End Function
Public Function IID_IAddressBarParser() As UUID
'{C9D81948-443A-40C7-945C-5E171B8C66B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC9D81948, CInt(&H443A), CInt(&H40C7), &H94, &H5C, &H5E, &H17, &H1B, &H8C, &H66, &HB4)
IID_IAddressBarParser = iid
End Function
Public Function IID_IWizardSite() As UUID
'{88960f5b-422f-4e7b-8013-73415381c3c3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88960F5B, CInt(&H422F), CInt(&H4E7B), &H80, &H13, &H73, &H41, &H53, &H81, &HC3, &HC3)
IID_IWizardSite = iid
End Function
Public Function IID_IWizardExtension() As UUID
'{c02ea696-86cc-491e-9b23-74394a0444a8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC02EA696, CInt(&H86CC), CInt(&H491E), &H9B, &H23, &H74, &H39, &H4A, &H4, &H44, &HA8)
IID_IWizardExtension = iid
End Function
Public Function IID_IFolderViewHost() As UUID
'{1ea58f02-d55a-411d-b09e-9e65ac21605b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1EA58F02, CInt(&HD55A), CInt(&H411D), &HB0, &H9E, &H9E, &H65, &HAC, &H21, &H60, &H5B)
IID_IFolderViewHost = iid
End Function
Public Function IID_IExtractIconA() As UUID
'{000214EB-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EB, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IExtractIconA = iid
End Function
Public Function IID_IExtractIconW() As UUID
'{000214FA-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214FA, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IExtractIconW = iid
End Function
Public Function IID_IShellPropSheetExt() As UUID
'{000214E9-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E9, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellPropSheetExt = iid
End Function
Public Function IID_IQueryInfo() As UUID
'{00021500-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21500, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IQueryInfo = iid
End Function
Public Function IID_IExtractImage2() As UUID
'{953BB1EE-93B4-11d1-98A3-00C04FB687DA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H953BB1EE, CInt(&H93B4), CInt(&H11D1), &H98, &HA3, &H0, &HC0, &H4F, &HB6, &H87, &HDA)
IID_IExtractImage2 = iid
End Function
Public Function IID_ICopyHookA() As UUID
'{000214EF-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214EF, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICopyHookA = iid
End Function
Public Function IID_ICopyHookW() As UUID
'{000214FC-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214FC, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ICopyHookW = iid
End Function
Public Function IID_IColumnProvider() As UUID
'{E8025004-1C42-11d2-BE2C-00A0C9A83DA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE8025004, CInt(&H1C42), CInt(&H11D2), &HBE, &H2C, &H0, &HA0, &HC9, &HA8, &H3D, &HA1)
IID_IColumnProvider = iid
End Function
Public Function IID_IURLSearchHook() As UUID
'{ac60f6a0-0fd9-11d0-99cb-00c04fd64497}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAC60F6A0, CInt(&HFD9), CInt(&H11D0), &H99, &HCB, &H0, &HC0, &H4F, &HD6, &H44, &H97)
IID_IURLSearchHook = iid
End Function
Public Function IID_ISearchContext() As UUID
'{09F656A2-41AF-480C-88F7-16CC0D164615}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F656A2, CInt(&H41AF), CInt(&H480C), &H88, &HF7, &H16, &HCC, &HD, &H16, &H46, &H15)
IID_ISearchContext = iid
End Function
Public Function IID_IURLSearchHook2() As UUID
'{5ee44da4-6d32-46e3-86bc-07540dedd0e0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5EE44DA4, CInt(&H6D32), CInt(&H46E3), &H86, &HBC, &H7, &H54, &HD, &HED, &HD0, &HE0)
IID_IURLSearchHook2 = iid
End Function
Public Function IID_INewShortcutHookA() As UUID
'{000214e1-0000-0000-c000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_INewShortcutHookA = iid
End Function
Public Function IID_INewShortcutHookW() As UUID
'{000214f7-0000-0000-c000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F7, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_INewShortcutHookW = iid
End Function
Public Function IID_ILayoutStorage() As UUID
'{0e6d4d90-6738-11cf-9608-00aa00680db4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE6D4D90, CInt(&H6738), CInt(&H11CF), &H96, &H8, &H0, &HAA, &H0, &H68, &HD, &HB4)
IID_ILayoutStorage = iid
End Function
Public Function IID_ISequentialStream() As UUID
'{0c733a30-2a1c-11ce-ade5-00aa0044773d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC733A30, CInt(&H2A1C), CInt(&H11CE), &HAD, &HE5, &H0, &HAA, &H0, &H44, &H77, &H3D)
IID_ISequentialStream = iid
End Function
Public Function IID_ITaskTrigger() As UUID
'{148BD52B-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD52B, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ITaskTrigger = iid
End Function
Public Function IID_IScheduledWorkItem() As UUID
'{a6b952f0-a4b1-11d0-997d-00aa006887ec}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6B952F0, CInt(&HA4B1), CInt(&H11D0), &H99, &H7D, &H0, &HAA, &H0, &H68, &H87, &HEC)
IID_IScheduledWorkItem = iid
End Function
Public Function IID_ITask() As UUID
'{148BD524-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD524, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ITask = iid
End Function
Public Function IID_IEnumWorkItems() As UUID
'{148BD528-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD528, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_IEnumWorkItems = iid
End Function
Public Function IID_ISchedulingAgent() As UUID
'{148BD527-A2AB-11CE-B11F-00AA00530503}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H148BD527, CInt(&HA2AB), CInt(&H11CE), &HB1, &H1F, &H0, &HAA, &H0, &H53, &H5, &H3)
IID_ISchedulingAgent = iid
End Function
Public Function IID_IResultsFolder() As UUID
'{96E5AE6D-6AE1-4b1c-900C-C6480EAA8828}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H96E5AE6D, CInt(&H6AE1), CInt(&H4B1C), &H90, &HC, &HC6, &H48, &HE, &HAA, &H88, &H28)
 IID_IResultsFolder = iid
End Function
Public Function IID_IVirtualDesktopManager() As UUID
'{a5cd92ff-29be-454c-8d04-d82879fb3f1b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA5CD92FF, CInt(&H29BE), CInt(&H454C), &H8D, &H4, &HD8, &H28, &H79, &HFB, &H3F, &H1B)
 IID_IVirtualDesktopManager = iid
End Function
Public Function IID_IInitializeNetworkFolder() As UUID
'{6e0f9881-42a8-4f2a-97f8-8af4e026d92d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E0F9881, CInt(&H42A8), CInt(&H4F2A), &H97, &HF8, &H8A, &HF4, &HE0, &H26, &HD9, &H2D)
 IID_IInitializeNetworkFolder = iid
End Function
Public Function IID_IProvideTaskPage() As UUID
'{4086658a-cbbb-11cf-b604-00c04fd8d565}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4086658A, CInt(&HCBBB), CInt(&H11CF), &HB6, &H4, &H0, &HC0, &H4F, &HD8, &HD5, &H65)
IID_IProvideTaskPage = iid
End Function
Public Function IID_ITextDocument() As UUID
'{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C0, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextDocument = iid
End Function
Public Function IID_ITextRange() As UUID
'{8CC497C2-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C2, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextRange = iid
End Function
Public Function IID_ITextSelection() As UUID
'{8CC497C1-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C1, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextSelection = iid
End Function
Public Function IID_ITextFont() As UUID
'{8CC497C3-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C3, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextFont = iid
End Function
Public Function IID_ITextPara() As UUID
'{8CC497C4-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C4, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextPara = iid
End Function
Public Function IID_ITextStoryRanges() As UUID
'{8CC497C5-A1DF-11CE-8098-00AA0047BE5D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CC497C5, CInt(&HA1DF), CInt(&H11CE), &H80, &H98, &H0, &HAA, &H0, &H47, &HBE, &H5D)
IID_ITextStoryRanges = iid
End Function
Public Function IID_ITypeInfo() As UUID
'{00020401-0000-0000-C000-000000000146}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20401, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H1, &H46)
IID_ITypeInfo = iid
End Function
Public Function IID_ITypeLib2() As UUID
'{00020411-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20411, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeLib2 = iid
End Function
Public Function IID_ITypeComp() As UUID
'{00020403-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20403, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_ITypeComp = iid
End Function
Public Function IID_IProvideClassInfo() As UUID
'{B196B283-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B283, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IProvideClassInfo = iid
End Function
Public Function IID_IConnectionPointContainer() As UUID
'{B196B284-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B284, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IConnectionPointContainer = iid
End Function
Public Function IID_IConnectionPoint() As UUID
'{B196B286-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B286, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IConnectionPoint = iid
End Function
Public Function IID_IDispatch() As UUID
'{00020400-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H20400, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IDispatch = iid
End Function
Public Function IID_IClassFactory() As UUID
'{00000001-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IClassFactory = iid
End Function
Public Function IID_IClassFactory2() As UUID
'{B196B28F-BAB4-101A-B69C-00AA00341D07}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB196B28F, CInt(&HBAB4), CInt(&H101A), &HB6, &H9C, &H0, &HAA, &H0, &H34, &H1D, &H7)
IID_IClassFactory2 = iid
End Function
Public Function IID_IUniformResourceLocatorA() As UUID
'{FBF23B80-E3F0-101B-8488-00AA003E56F8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFBF23B80, CInt(&HE3F0), CInt(&H101B), &H84, &H88, &H0, &HAA, &H0, &H3E, &H56, &HF8)
IID_IUniformResourceLocatorA = iid
End Function
Public Function IID_IUniformResourceLocatorW() As UUID
'{CABB0DA0-DA57-11CF-9974-0020AFD79762}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCABB0DA0, CInt(&HDA57), CInt(&H11CF), &H99, &H74, &H0, &H20, &HAF, &HD7, &H97, &H62)
IID_IUniformResourceLocatorW = iid
End Function
Public Function IID_IEnumSTATURL() As UUID
'{3C374A42-BAE4-11CF-BF7D-00AA006946EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C374A42, CInt(&HBAE4), CInt(&H11CF), &HBF, &H7D, &H0, &HAA, &H0, &H69, &H46, &HEE)
IID_IEnumSTATURL = iid
End Function
Public Function IID_IUrlHistoryStg() As UUID
'{3C374A41-BAE4-11CF-BF7D-00AA006946EE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3C374A41, CInt(&HBAE4), CInt(&H11CF), &HBF, &H7D, &H0, &HAA, &H0, &H69, &H46, &HEE)
IID_IUrlHistoryStg = iid
End Function
Public Function IID_IUrlHistoryStg2() As UUID
'{AFA0DC11-C313-11d0-831A-00C04FD5AE38}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAFA0DC11, CInt(&HC313), CInt(&H11D0), &H83, &H1A, &H0, &HC0, &H4F, &HD5, &HAE, &H38)
IID_IUrlHistoryStg2 = iid
End Function
Public Function IID_IBinding() As UUID
'{79eac9c0-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C0, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBinding = iid
End Function
Public Function IID_IBindStatusCallback() As UUID
'{79eac9c1-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9C1, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBindStatusCallback = iid
End Function
Public Function IID_IAuthenticate() As UUID
'{79eac9d0-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D0, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IAuthenticate = iid
End Function
Public Function IID_IInternetProtocolInfo() As UUID
'{79eac9ec-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EC, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolInfo = iid
End Function
Public Function IID_IInternetPriority() As UUID
'{79eac9eb-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9EB, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetPriority = iid
End Function
Public Function IID_IInternetSession() As UUID
'{79eac9e7-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E7, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetSession = iid
End Function
Public Function IID_IInternetProtocolRoot() As UUID
'{79eac9e3-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E3, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolRoot = iid
End Function
Public Function IID_IInternetProtocol() As UUID
'{79eac9e4-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E4, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocol = iid
End Function
Public Function IID_IInternetProtocolSink() As UUID
'{79eac9e5-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E5, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetProtocolSink = iid
End Function
Public Function IID_IInternetBindInfo() As UUID
'{79eac9e1-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9E1, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IInternetBindInfo = iid
End Function
Public Function IID_IBindProtocol() As UUID
'{79eac9cd-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9CD, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IBindProtocol = iid
End Function
Public Function IID_IHttpNegotiate() As UUID
'{79eac9d2-baf9-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D2, CInt(&HBAF9), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHttpNegotiate = iid
End Function
Public Function IID_IWindowForBindingUI() As UUID
'{79eac9d5-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D5, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWindowForBindingUI = iid
End Function
Public Function IID_IWinInetInfo() As UUID
'{79eac9d6-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D6, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWinInetInfo = iid
End Function
Public Function IID_IWinInetHttpInfo() As UUID
'{79eac9d8-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D8, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IWinInetHttpInfo = iid
End Function
Public Function IID_IBindHost() As UUID
'{fc4801a1-2ba9-11cf-a229-00aa003d7352}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFC4801A1, CInt(&H2BA9), CInt(&H11CF), &HA2, &H29, &H0, &HAA, &H0, &H3D, &H73, &H52)
IID_IBindHost = iid
End Function
Public Function IID_IHttpNegotiate2() As UUID
'{4F9F9FCB-E0F4-48eb-B7AB-FA2EA9365CB4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4F9F9FCB, CInt(&HE0F4), CInt(&H48EB), &HB7, &HAB, &HFA, &H2E, &HA9, &H36, &H5C, &HB4)
IID_IHttpNegotiate2 = iid
End Function
Public Function IID_IHttpSecurity() As UUID
'{79eac9d7-bafa-11ce-8c82-00aa004ba90b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79EAC9D7, CInt(&HBAFA), CInt(&H11CE), &H8C, &H82, &H0, &HAA, &H0, &H4B, &HA9, &HB)
IID_IHttpSecurity = iid
End Function
Public Function IID_IViewObject() As UUID
'{0000010D-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IViewObject = iid
End Function
Public Function IID_IViewObject2() As UUID
'{00000127-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H127, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IViewObject2 = iid
End Function
Public Function IID_IWMPRemoteMediaServices() As UUID
'{CBB92747-741F-44fe-AB5B-F1A48F3B2A59}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCBB92747, CInt(&H741F), CInt(&H44FE), &HAB, &H5B, &HF1, &HA4, &H8F, &H3B, &H2A, &H59)
IID_IWMPRemoteMediaServices = iid
End Function
Public Function IID_IWMPPluginUI() As UUID
'{4C5E8F9F-AD3E-4bf9-9753-FCD30D6D38DD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C5E8F9F, CInt(&HAD3E), CInt(&H4BF9), &H97, &H53, &HFC, &HD3, &HD, &H6D, &H38, &HDD)
IID_IWMPPluginUI = iid
End Function
'IID_IShellView =    { 0x000214E3, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
Public Function IID_IShellView() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E3, 0, 0)
 IID_IShellView = iid
End Function
Public Function IID_IFolderView() As UUID
'{cde725b0-ccc9-4519-917e-325d72fab4ce}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDE725B0, CInt(&HCCC9), CInt(&H4519), &H91, &H7E, &H32, &H5D, &H72, &HFA, &HB4, &HCE)
 IID_IFolderView = iid
End Function
Public Function IID_IFolderView2() As UUID
'{1af3a467-214f-4298-908e-06b03e0b39f9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AF3A467, CInt(&H214F), CInt(&H4298), &H90, &H8E, &H6, &HB0, &H3E, &HB, &H39, &HF9)
 IID_IFolderView2 = iid
End Function

' Returns the IShellFolder interface ID, {000214E6-0000-0000-C000-000000046}
Public Function IID_IShellFolder() As UUID
  Static iid As UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E6, 0, 0)
  IID_IShellFolder = iid
End Function

' Returns the IShellDetails interface ID,

Public Function IID_IShellDetails() As UUID
'{000214EC-0000-0000-C000-000000000046}
  Static iid As UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214EC, 0, 0)
  IID_IShellDetails = iid
End Function
Public Function IID_IExtractImage() As UUID
'{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBB2E617C, CInt(&H920), CInt(&H11D1), &H9A, &HB, &H0, &HC0, &H4F, &HC2, &HD6, &HC1)
  IID_IExtractImage = iid

End Function
Public Function IID_IShellFolder2() As UUID
'{93F2F68C-1D1B-11D3-A30E-00C04F79ABD1}
    Static iid As UUID
    If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H93F2F68C, CInt(&H1D1B), CInt(&H11D3), &HA3, &HE, 0, &HC0, &H4F, &H79, &HAB, &HD1)
    IID_IShellFolder2 = iid
End Function

Public Function IID_IStorage() As UUID
'({0000000B-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &HB, 0, 0)
 IID_IStorage = iid
End Function
Public Function IID_IRootStorage() As UUID
'({00000012-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H12, 0, 0)
 IID_IRootStorage = iid
End Function
Public Function IID_IPropertyStorage() As UUID
'({00000138-0000-0000-C000-000000000046})
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H12, 0, 0)
 IID_IPropertyStorage = iid
End Function
Public Function IID_IShellItem() As UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Public Function IID_IShellItem2() As UUID
'7e9fb0d3-919f-4307-ab2e-9b1860310c93
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E9FB0D3, CInt(&H919F), CInt(&H4307), &HAB, &H2E, &H9B, &H18, &H60, &H31, &HC, &H93)
IID_IShellItem2 = iid
End Function
Public Function IID_IEnumShellItems() As UUID
'{70629033-e363-4a28-a567-0db78006e6d7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70629033, CInt(&HE363), CInt(&H4A28), &HA5, &H67, &HD, &HB7, &H80, &H6, &HE6, &HD7)
 IID_IEnumShellItems = iid
End Function
Public Function IID_IShellLibrary() As UUID
'{0x11a66efa, 0x382e, 0x451a, {0x92, 0x34, 0x1e, 0xe, 0x12, 0xef, 0x30, 0x85}}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11A66EFA, CInt(&H382E), CInt(&H451A), &H92, &H34, &H1E, &HE, &H12, &HEF, &H30, &H85)
  IID_IShellLibrary = iid

End Function
Public Function IID_IShellItemArray() As UUID
'{b63ea76d-1f85-456f-a19c-48159efa858b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB63EA76D, CInt(&H1F85), CInt(&H456F), &HA1, &H9C, &H48, &H15, &H9E, &HFA, &H85, &H8B)
  IID_IShellItemArray = iid

End Function
Public Function IID_IObjectArray() As UUID
'0x92ca9dcd, 0x5622, 0x4bba, 0xa8,0x05, 0x5e,0x9f,0x54,0x1b,0xd8,0xc9
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H92CA9DCD, CInt(&H5622), CInt(&H4BBA), &HA8, &H5, &H5E, &H9F, &H54, &H1B, &HD8, &HC9)
  IID_IObjectArray = iid

End Function
Public Function IID_IShellItemImageFactory() As UUID
'{BCC18B79-BA16-442F-80C4-8A59C30C463B}
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCC18B79, CInt(&HBA16), CInt(&H442F), &H80, &HC4, &H8A, &H59, &HC3, &HC, &H46, &H3B)
IID_IShellItemImageFactory = iid
End Function
Public Function IID_IOleLink() As UUID
'{0000011d-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11D, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IOleLink = iid
End Function
Public Function IID_IPropertySetStorage() As UUID
'{0000013A-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H13A, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IPropertySetStorage = iid
End Function
Public Function IID_ICondition() As UUID
'{0FC988D4-C935-4b97-A973-46282EA175C8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFC988D4, CInt(&HC935), CInt(&H4B97), &HA9, &H73, &H46, &H28, &H2E, &HA1, &H75, &HC8)
 IID_ICondition = iid
End Function

Public Function IID_IDataObject() As UUID
'0000010e-0000-0000-C000-000000000046
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
  IID_IDataObject = iid

End Function

Public Function IID_IFileDialogCustomize() As UUID
'IID_IFileDialogCustomize "{8016b7b3-3d49-4504-a0aa-2a37494e606f}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8016B7B3, CInt(&H3D49), CInt(&H4504), &HA0, &HAA, &H2A, &H37, &H49, &H4E, &H60, &H6F)
  IID_IFileDialogCustomize = iid

End Function
Public Function IID_IShellMenu() As UUID
'{0x1FEAEBFA,0x3C7A,0x4BB6,{0xB0,0xD2,0xF1,0xB8,0x1B,0x8F,0x27,0xED}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1FEAEBFA, CInt(&H3C7A), CInt(&H4BB6), &HB0, &HD2, &HF1, &HB8, &H1B, &H8F, &H27, &HED)
  IID_IShellMenu = iid
  
End Function
Public Function IID_IPropertyDescriptionList() As UUID
'IID_IPropertyDescriptionList, 0x1f9fc1d0, 0xc39b, 0x4b26, 0x81,0x7f, 0x01,0x19,0x67,0xd3,0x44,0x0e
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F9FC1D0, CInt(&HC39B), CInt(&H4B26), &H81, &H7F, &H1, &H19, &H67, &HD3, &H44, &HE)
  IID_IPropertyDescriptionList = iid

End Function

Public Function IID_IPropertyDescription() As UUID
'(IID_IPropertyDescription, 0x6f79d558, 0x3e96, 0x4549, 0xa1,0xd1, 0x7d,0x75,0xd2,0x28,0x88,0x14
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F79D558, CInt(&H3E96), CInt(&H4549), &HA1, &HD1, &H7D, &H75, &HD2, &H28, &H88, &H14)
  IID_IPropertyDescription = iid
  
End Function

Public Function IID_IPropertyStore() As UUID
'DEFINE_GUID(IID_IPropertyStore,0x886d8eeb, 0x8cf2, 0x4446, 0x8d,0x02,0xcd,0xba,0x1d,0xbd,0xcf,0x99);
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H886D8EEB, CInt(&H8CF2), CInt(&H4446), &H8D, &H2, &HCD, &HBA, &H1D, &HBD, &HCF, &H99)
  IID_IPropertyStore = iid
  
End Function
Public Function IID_IPropertySystem() As UUID
'IID_IPropertySystem, 0xca724e8a, 0xc3e6, 0x442b, 0x88,0xa4, 0x6f,0xb0,0xdb,0x80,0x35,0xa3
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCA724E8A, CInt(&HC3E6), CInt(&H442B), &H88, &HA4, &H6F, &HB0, &HDB, &H80, &H35, &HA3)
  IID_IPropertySystem = iid
  
End Function
Public Function IID_IDropTarget() As UUID
'{00000122-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H122, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IDropTarget = iid
End Function
Public Function IID_IDropSource() As UUID
'{00000121-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H121, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IDropSource = iid
End Function
Public Function IID_IDragSourceHelper() As UUID
'{de5bf786-477a-11d2-839d-00c04fd918d0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE5BF786, CInt(&H477A), CInt(&H11D2), &H83, &H9D, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
  IID_IDragSourceHelper = iid
  
End Function

Public Function IID_IDragSourceHelper2() As UUID
'{83E07D0D-0C5F-4163-BF1A-60B274051E40}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H83E07D0D, CInt(&HC5F), CInt(&H4163), &HBF, &H1A, &H60, &HB2, &H74, &H5, &H1E, &H40)
  IID_IDragSourceHelper2 = iid
  
End Function
Public Function IID_IDropTargetHelper() As UUID
'{4657278B-411B-11D2-839A-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4657278B, CInt(&H411B), CInt(&H11D2), &H83, &H9A, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
 IID_IDropTargetHelper = iid
End Function

Public Function CLSID_QueryAssociations() As UUID
'{a07034fd-6caa-4954-ac3f-97a27216f98a}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA07034FD, CInt(&H6CAA), CInt(&H4954), &HAC, &H3F, &H97, &HA2, &H72, &H16, &HF9, &H8A)
 CLSID_QueryAssociations = iid
End Function

Public Function CLSID_ImageList() As UUID
'{7C476BA2-02B1-48f4-8048-B24619DDC058}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C476BA2, CInt(&H2B1), CInt(&H48F4), &H80, &H48, &HB2, &H46, &H19, &HDD, &HC0, &H58)
 CLSID_ImageList = iid
End Function

Public Function IID_IQueryAssociations() As UUID
'{c46ca590-3c3f-11d2-bee6-0000f805ca57}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC46CA590, CInt(&H3C3F), CInt(&H11D2), &HBE, &HE6, &H0, &H0, &HF8, &H5, &HCA, &H57)
 IID_IQueryAssociations = iid
End Function

Public Function IID_IPreviewHandler() As UUID
'{8895b1c6-b41f-4c1c-a562-0d564250836f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8895B1C6, CInt(&HB41F), CInt(&H4C1C), &HA5, &H62, &HD, &H56, &H42, &H50, &H83, &H6F)
 IID_IPreviewHandler = iid
End Function
Public Function IID_IPreviewHandlerVisuals() As UUID
'{196bf9a5-b346-4ef0-aa1e-5dcdb76768b1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H196BF9A5, CInt(&HB346), CInt(&H4EF0), &HAA, &H1E, &H5D, &HCD, &HB7, &H67, &H68, &HB1)
 IID_IPreviewHandlerVisuals = iid
End Function
Public Function IID_IInitializeWithStream() As UUID
'{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB824B49D, CInt(&H22AC), CInt(&H4161), &HAC, &H8A, &H99, &H16, &HE8, &HFA, &H3F, &H7F)
 IID_IInitializeWithStream = iid
End Function
Public Function IID_IInitializeWithFile() As UUID
'{b7d14566-0509-4cce-a71f-0a554233bd9b}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7D14566, CInt(&H509), CInt(&H4CCE), &HA7, &H1F, &HA, &H55, &H42, &H33, &HBD, &H9B)
 IID_IInitializeWithFile = iid
End Function
Public Function IID_IInitializeWithItem() As UUID
'{7f73be3f-fb79-493c-a6c7-7ee14e245841}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F73BE3F, CInt(&HFB79), CInt(&H493C), &HA6, &HC7, &H7E, &HE1, &H4E, &H24, &H58, &H41)
 IID_IInitializeWithItem = iid
End Function
Public Function IID_IInitializeWithPropertyStore() As UUID
'{C3E12EB5-7D8D-44f8-B6DD-0E77B34D6DE4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC3E12EB5, CInt(&H7D8D), CInt(&H44F8), &HB6, &HDD, &HE, &H77, &HB3, &H4D, &H6D, &HE4)
 IID_IInitializeWithPropertyStore = iid
End Function
Public Function IID_IInitializeWithWindow() As UUID
'{3E68D4BD-7135-4D10-8018-9FB6D9F33FA1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3E68D4BD, CInt(&H7135), CInt(&H4D10), &H80, &H18, &H9F, &HB6, &HD9, &HF3, &H3F, &HA1)
 IID_IInitializeWithWindow = iid
End Function
Public Function IID_ICreateObject() As UUID
'{75121952-e0d0-43e5-9380-1d80483acf72}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H75121952, CInt(&HE0D0), CInt(&H43E5), &H93, &H80, &H1D, &H80, &H48, &H3A, &HCF, &H72)
 IID_ICreateObject = iid
End Function

Public Function IID_IPropertyBag() As UUID
'{55272A00-42CB-11CE-8135-00AA004BB851}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H55272A00, CInt(&H42CB), CInt(&H11CE), &H81, &H35, &H0, &HAA, &H0, &H4B, &HB8, &H51)
 IID_IPropertyBag = iid
End Function

Public Function IID_IImageList() As UUID
'{46EB5926-582E-4017-9FDF-E8998DAA0950}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46EB5926, CInt(&H582E), CInt(&H4017), &H9F, &HDF, &HE8, &H99, &H8D, &HAA, &H9, &H50)
 IID_IImageList = iid
End Function
Public Function IID_IImageList2() As UUID
'{192b9d83-50fc-457b-90a0-2b82a8b5dae1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H192B9D83, CInt(&H50FC), CInt(&H457B), &H90, &HA0, &H2B, &H82, &HA8, &HB5, &HDA, &HE1)
 IID_IImageList2 = iid
End Function
Public Function IID_IContextMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E4, 0, 0)
 IID_IContextMenu = iid
End Function
Public Function IID_IContextMenu2() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214F4, 0, 0)
 IID_IContextMenu2 = iid
End Function
Public Function IID_IContextMenu3() As UUID
'{BCFCE0A0-EC17-11d0-8D10-00A0C90F2719}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCFCE0A0, CInt(&HEC17), CInt(&H11D0), &H8D, &H10, &H0, &HA0, &HC9, &HF, &H27, &H19)
 IID_IContextMenu3 = iid
End Function
Public Function IID_IContextMenuCB() As UUID
'{3409E930-5A39-11d1-83FA-00A0C90DC849}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3409E930, CInt(&H5A39), CInt(&H11D1), &H83, &HFA, &H0, &HA0, &HC9, &HD, &HC8, &H49)
 IID_IContextMenuCB = iid
End Function
Public Function IID_IContextMenuSite() As UUID
'{0811AEBE-0B87-4C54-9E72-548CF649016B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H811AEBE, CInt(&HB87), CInt(&H4C54), &H9E, &H72, &H54, &H8C, &HF6, &H49, &H1, &H6B)
 IID_IContextMenuSite = iid
End Function
Public Function IID_IHomeGroup() As UUID
'{7a3bd1d9-35a9-4fb3-a467-f48cac35e2d0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7A3BD1D9, CInt(&H35A9), CInt(&H4FB3), &HA4, &H67, &HF4, &H8C, &HAC, &H35, &HE2, &HD0)
 IID_IHomeGroup = iid
End Function
Public Function IID_ICallQI() As UUID
'{9fb58518-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58518, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallQI = iid
End Function
Public Function IID_IMultiQI() As UUID
'{00000020-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H20, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IMultiQI = iid
End Function
Public Function IID_ICallAddRelease() As UUID
'{9fb58519-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58519, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallAddRelease = iid
End Function
Public Function IID_ICallGION() As UUID
'{9fb58520-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58520, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallGION = iid
End Function
Public Function IID_ICallInvoke() As UUID
'{9fb58521-92ec-4bf6-bc61-ff4e59df7369}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9FB58521, CInt(&H92EC), CInt(&H4BF6), &HBC, &H61, &HFF, &H4E, &H59, &HDF, &H73, &H69)
 IID_ICallInvoke = iid
End Function
Public Function IID_IDefaultExtractIconInit() As UUID
'{41ded17d-d6b3-4261-997d-88c60e4b1d58}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H41DED17D, CInt(&HD6B3), CInt(&H4261), &H99, &H7D, &H88, &HC6, &HE, &H4B, &H1D, &H58)
 IID_IDefaultExtractIconInit = iid
End Function
Public Function IID_IExecuteCommand() As UUID
'{7F9185B0-CB92-43c5-80A9-92277A4F7B54}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F9185B0, CInt(&HCB92), CInt(&H43C5), &H80, &HA9, &H92, &H27, &H7A, &H4F, &H7B, &H54)
 IID_IExecuteCommand = iid
End Function
Public Function IID_IExecuteCommandHost() As UUID
'{4b6832a2-5f04-4c9d-b89d-727a15d103e7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4B6832A2, CInt(&H5F04), CInt(&H4C9D), &HB8, &H9D, &H72, &H7A, &H15, &HD1, &H3, &HE7)
 IID_IExecuteCommandHost = iid
End Function
Public Function IID_IExplorerCommandProvider() As UUID
'{64961751-0835-43c0-8ffe-d57686530e64}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64961751, CInt(&H835), CInt(&H43C0), &H8F, &HFE, &HD5, &H76, &H86, &H53, &HE, &H64)
 IID_IExplorerCommandProvider = iid
End Function
Public Function IID_IEnumExplorerCommand() As UUID
'{a88826f8-186f-4987-aade-ea0cef8fbfe8}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA88826F8, CInt(&H186F), CInt(&H4987), &HAA, &HDE, &HEA, &HC, &HEF, &H8F, &HBF, &HE8)
 IID_IEnumExplorerCommand = iid
End Function
Public Function IID_IInitializeCommand() As UUID
'{85075acf-231f-40ea-9610-d26b7b58f638}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H85075ACF, CInt(&H231F), CInt(&H40EA), &H96, &H10, &HD2, &H6B, &H7B, &H58, &HF6, &H38)
 IID_IInitializeCommand = iid
End Function
Public Function IID_IExplorerCommandState() As UUID
'{bddacb60-7657-47ae-8445-d23e1acf82ae}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBDDACB60, CInt(&H7657), CInt(&H47AE), &H84, &H45, &HD2, &H3E, &H1A, &HCF, &H82, &HAE)
 IID_IExplorerCommandState = iid
End Function
Public Function IID_IExplorerCommand() As UUID
'{a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA08CE4D0, CInt(&HFA25), CInt(&H44AB), &HB5, &H7C, &HC7, &HB1, &HC3, &H23, &HE0, &HB9)
 IID_IExplorerCommand = iid
End Function
Public Function IID_IMessageFilter() As UUID
'{00000016-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H16, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IMessageFilter = iid
End Function
Public Function IID_IApplicationDocumentLists() As UUID
'{3c594f9f-9f30-47a1-979a-c9e83d3d0a06}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C594F9F, CInt(&H9F30), CInt(&H47A1), &H97, &H9A, &HC9, &HE8, &H3D, &H3D, &HA, &H6)
 IID_IApplicationDocumentLists = iid
End Function
Public Function IID_IShellChangeNotify() As UUID
'{D82BE2B1-5764-11D0-A96E-00C04FD705A2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD82BE2B1, CInt(&H5764), CInt(&H11D0), &HA9, &H6E, &H0, &HC0, &H4F, &HD7, &H5, &HA2)
 IID_IShellChangeNotify = iid
End Function
Public Function IID_ITransferSource() As UUID
'{00adb003-bde9-45c6-8e29-d09f9353e108}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HADB003, CInt(&HBDE9), CInt(&H45C6), &H8E, &H29, &HD0, &H9F, &H93, &H53, &HE1, &H8)
IID_ITransferSource = iid
End Function
Public Function IID_IEnumResources() As UUID
'{2dd81fe3-a83c-4da9-a330-47249d345ba1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2DD81FE3, CInt(&HA83C), CInt(&H4DA9), &HA3, &H30, &H47, &H24, &H9D, &H34, &H5B, &HA1)
IID_IEnumResources = iid
End Function
Public Function IID_IShellItemResources() As UUID
'{ff5693be-2ce0-4d48-b5c5-40817d1acdb9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFF5693BE, CInt(&H2CE0), CInt(&H4D48), &HB5, &HC5, &H40, &H81, &H7D, &H1A, &HCD, &HB9)
IID_IShellItemResources = iid
End Function
Public Function IID_ITransferDestination() As UUID
'{48addd32-3ca5-4124-abe3-b5a72531b207}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H48ADDD32, CInt(&H3CA5), CInt(&H4124), &HAB, &HE3, &HB5, &HA7, &H25, &H31, &HB2, &H7)
IID_ITransferDestination = iid
End Function
Public Function IID_IKnownFolder() As UUID
'{3AA7AF7E-9B36-420c-A8E3-F77D4674A488}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3AA7AF7E, CInt(&H9B36), CInt(&H420C), &HA8, &HE3, &HF7, &H7D, &H46, &H74, &HA4, &H88)
IID_IKnownFolder = iid
End Function
Public Function IID_IKnownFolderManager() As UUID
'{8BE2D872-86AA-4d47-B776-32CCA40C7018}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8BE2D872, CInt(&H86AA), CInt(&H4D47), &HB7, &H76, &H32, &HCC, &HA4, &HC, &H70, &H18)
IID_IKnownFolderManager = iid
End Function
Public Function IID_IInitializeWithBindCtx() As UUID
'{71c0d2bc-726d-45cc-a6c0-2e31c1db2159}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71C0D2BC, CInt(&H726D), CInt(&H45CC), &HA6, &HC0, &H2E, &H31, &HC1, &HDB, &H21, &H59)
IID_IInitializeWithBindCtx = iid
End Function
Public Function IID_IPreviewHandlerFrame() As UUID
'{fec87aaf-35f9-447a-adb7-20234491401a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFEC87AAF, CInt(&H35F9), CInt(&H447A), &HAD, &HB7, &H20, &H23, &H44, &H91, &H40, &H1A)
IID_IPreviewHandlerFrame = iid
End Function
Public Function IID_IVisualProperties() As UUID
'{e693cf68-d967-4112-8763-99172aee5e5a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE693CF68, CInt(&HD967), CInt(&H4112), &H87, &H63, &H99, &H17, &H2A, &HEE, &H5E, &H5A)
IID_IVisualProperties = iid
End Function
Public Function IID_ISpellingError() As UUID
'{B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7C82D61, CInt(&HFBE8), CInt(&H4B47), &H9B, &H27, &H6C, &HD, &H2E, &HD, &HE0, &HA3)
 IID_ISpellingError = iid
End Function
Public Function IID_IEnumSpellingError() As UUID
'{803E3BD4-2828-4410-8290-418D1D73C762}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H803E3BD4, CInt(&H2828), CInt(&H4410), &H82, &H90, &H41, &H8D, &H1D, &H73, &HC7, &H62)
 IID_IEnumSpellingError = iid
End Function
Public Function IID_IOptionDescription() As UUID
'{432E5F85-35CF-4606-A801-6F70277E1D7A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H432E5F85, CInt(&H35CF), CInt(&H4606), &HA8, &H1, &H6F, &H70, &H27, &H7E, &H1D, &H7A)
 IID_IOptionDescription = iid
End Function
Public Function IID_ISpellCheckerChangedEventHandler() As UUID
'{0B83A5B0-792F-4EAB-9799-ACF52C5ED08A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB83A5B0, CInt(&H792F), CInt(&H4EAB), &H97, &H99, &HAC, &HF5, &H2C, &H5E, &HD0, &H8A)
 IID_ISpellCheckerChangedEventHandler = iid
End Function
Public Function IID_ISpellChecker() As UUID
'{B6FD0B71-E2BC-4653-8D05-F197E412770B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB6FD0B71, CInt(&HE2BC), CInt(&H4653), &H8D, &H5, &HF1, &H97, &HE4, &H12, &H77, &HB)
 IID_ISpellChecker = iid
End Function
Public Function IID_ISpellChecker2() As UUID
'{E7ED1C71-87F7-4378-A840-C9200DACEE47}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE7ED1C71, CInt(&H87F7), CInt(&H4378), &HA8, &H40, &HC9, &H20, &HD, &HAC, &HEE, &H47)
 IID_ISpellChecker2 = iid
End Function
Public Function IID_ISpellCheckerFactory() As UUID
'{8E018A9D-2415-4677-BF08-794EA61F94BB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8E018A9D, CInt(&H2415), CInt(&H4677), &HBF, &H8, &H79, &H4E, &HA6, &H1F, &H94, &HBB)
 IID_ISpellCheckerFactory = iid
End Function
Public Function IID_IUserDictionariesRegistrar() As UUID
'{AA176B85-0E12-4844-8E1A-EEF1DA77F586}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAA176B85, CInt(&HE12), CInt(&H4844), &H8E, &H1A, &HEE, &HF1, &HDA, &H77, &HF5, &H86)
 IID_IUserDictionariesRegistrar = iid
End Function
Public Function IID_ISpellCheckProvider() As UUID
'{73E976E0-8ED4-4EB1-80D7-1BE0A16B0C38}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73E976E0, CInt(&H8ED4), CInt(&H4EB1), &H80, &HD7, &H1B, &HE0, &HA1, &H6B, &HC, &H38)
 IID_ISpellCheckProvider = iid
End Function
Public Function IID_IComprehensiveSpellCheckProvider() As UUID
'{0C58F8DE-8E94-479E-9717-70C42C4AD2C3}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC58F8DE, CInt(&H8E94), CInt(&H479E), &H97, &H17, &H70, &HC4, &H2C, &H4A, &HD2, &HC3)
 IID_IComprehensiveSpellCheckProvider = iid
End Function
Public Function IID_ISpellCheckProviderFactory() As UUID
'{9F671E11-77D6-4C92-AEFB-615215E3A4BE}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9F671E11, CInt(&H77D6), CInt(&H4C92), &HAE, &HFB, &H61, &H52, &H15, &HE3, &HA4, &HBE)
 IID_ISpellCheckProviderFactory = iid
End Function
Public Function IID_IRichChunk() As UUID
'{4FDEF69C-DBC9-454e-9910-B34F3C64B510}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4FDEF69C, CInt(&HDBC9), CInt(&H454E), &H99, &H10, &HB3, &H4F, &H3C, &H64, &HB5, &H10)
IID_IRichChunk = iid
End Function
Public Function IID_ICondition2() As UUID
'{0DB8851D-2E5B-47eb-9208-D28C325A01D7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDB8851D, CInt(&H2E5B), CInt(&H47EB), &H92, &H8, &HD2, &H8C, &H32, &H5A, &H1, &HD7)
IID_ICondition2 = iid
End Function
Public Function IID_IConditionFactory() As UUID
'{A5EFE073-B16F-474f-9F3E-9F8B497A3E08}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA5EFE073, CInt(&HB16F), CInt(&H474F), &H9F, &H3E, &H9F, &H8B, &H49, &H7A, &H3E, &H8)
IID_IConditionFactory = iid
End Function
Public Function IID_IConditionFactory2() As UUID
'{71D222E1-432F-429e-8C13-B6DAFDE5077A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H71D222E1, CInt(&H432F), CInt(&H429E), &H8C, &H13, &HB6, &HDA, &HFD, &HE5, &H7, &H7A)
IID_IConditionFactory2 = iid
End Function
Public Function IID_IQueryParser() As UUID
'{2EBDEE67-3505-43f8-9946-EA44ABC8E5B0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2EBDEE67, CInt(&H3505), CInt(&H43F8), &H99, &H46, &HEA, &H44, &HAB, &HC8, &HE5, &HB0)
IID_IQueryParser = iid
End Function
Public Function IID_IQuerySolution() As UUID
'{D6EBC66B-8921-4193-AFDD-A1789FB7FF57}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD6EBC66B, CInt(&H8921), CInt(&H4193), &HAF, &HDD, &HA1, &H78, &H9F, &HB7, &HFF, &H57)
IID_IQuerySolution = iid
End Function
Public Function IID_IConditionGenerator() As UUID
'{92D2CC58-4386-45a3-B98C-7E0CE64A4117}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H92D2CC58, CInt(&H4386), CInt(&H45A3), &HB9, &H8C, &H7E, &HC, &HE6, &H4A, &H41, &H17)
IID_IConditionGenerator = iid
End Function
Public Function IID_IInterval() As UUID
'{6BF0A714-3C18-430b-8B5D-83B1C234D3DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6BF0A714, CInt(&H3C18), CInt(&H430B), &H8B, &H5D, &H83, &HB1, &HC2, &H34, &HD3, &HDB)
IID_IInterval = iid
End Function
Public Function IID_IMetaData() As UUID
'{780102B0-C43B-4876-BC7B-5E9BA5C88794}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H780102B0, CInt(&HC43B), CInt(&H4876), &HBC, &H7B, &H5E, &H9B, &HA5, &HC8, &H87, &H94)
IID_IMetaData = iid
End Function
Public Function IID_IEntity() As UUID
'{24264891-E80B-4fd3-B7CE-4FF2FAE8931F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24264891, CInt(&HE80B), CInt(&H4FD3), &HB7, &HCE, &H4F, &HF2, &HFA, &HE8, &H93, &H1F)
IID_IEntity = iid
End Function
Public Function IID_IRelationship() As UUID
'{2769280B-5108-498c-9C7F-A51239B63147}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2769280B, CInt(&H5108), CInt(&H498C), &H9C, &H7F, &HA5, &H12, &H39, &HB6, &H31, &H47)
IID_IRelationship = iid
End Function
Public Function IID_INamedEntity() As UUID
'{ABDBD0B1-7D54-49fb-AB5C-BFF4130004CD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HABDBD0B1, CInt(&H7D54), CInt(&H49FB), &HAB, &H5C, &HBF, &HF4, &H13, &H0, &H4, &HCD)
IID_INamedEntity = iid
End Function
Public Function IID_ISchemaProvider() As UUID
'{8CF89BCB-394C-49b2-AE28-A59DD4ED7F68}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CF89BCB, CInt(&H394C), CInt(&H49B2), &HAE, &H28, &HA5, &H9D, &HD4, &HED, &H7F, &H68)
IID_ISchemaProvider = iid
End Function
Public Function IID_ITokenCollection() As UUID
'{22D8B4F2-F577-4adb-A335-C2AE88416FAB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22D8B4F2, CInt(&HF577), CInt(&H4ADB), &HA3, &H35, &HC2, &HAE, &H88, &H41, &H6F, &HAB)
IID_ITokenCollection = iid
End Function
Public Function IID_INamedEntityCollector() As UUID
'{AF2440F6-8AFC-47d0-9A7F-396A0ACFB43D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAF2440F6, CInt(&H8AFC), CInt(&H47D0), &H9A, &H7F, &H39, &H6A, &HA, &HCF, &HB4, &H3D)
IID_INamedEntityCollector = iid
End Function
Public Function IID_ISchemaLocalizerSupport() As UUID
'{CA3FDCA2-BFBE-4eed-90D7-0CAEF0A1BDA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCA3FDCA2, CInt(&HBFBE), CInt(&H4EED), &H90, &HD7, &HC, &HAE, &HF0, &HA1, &HBD, &HA1)
IID_ISchemaLocalizerSupport = iid
End Function
Public Function IID_IQueryParserManager() As UUID
'{A879E3C4-AF77-44fb-8F37-EBD1487CF920}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA879E3C4, CInt(&HAF77), CInt(&H44FB), &H8F, &H37, &HEB, &HD1, &H48, &H7C, &HF9, &H20)
IID_IQueryParserManager = iid
End Function
Public Function IID_ISearchFolderItemFactory() As UUID
'{a0ffbc28-5482-4366-be27-3e81e78e06c2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0FFBC28, CInt(&H5482), CInt(&H4366), &HBE, &H27, &H3E, &H81, &HE7, &H8E, &H6, &HC2)
IID_ISearchFolderItemFactory = iid
End Function
Public Function IID_IThumbnailHandlerFactory() As UUID
'{e35b4b2e-00da-4bc1-9f13-38bc11f5d417}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE35B4B2E, CInt(&HDA), CInt(&H4BC1), &H9F, &H13, &H38, &HBC, &H11, &HF5, &HD4, &H17)
IID_IThumbnailHandlerFactory = iid
End Function
Public Function IID_ISharedBitmap() As UUID
'{091162a4-bc96-411f-aae8-c5122cd03363}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H91162A4, CInt(&HBC96), CInt(&H411F), &HAA, &HE8, &HC5, &H12, &H2C, &HD0, &H33, &H63)
 IID_ISharedBitmap = iid
End Function
Public Function IID_IThumbnailCache() As UUID
'{F676C15D-596A-4ce2-8234-33996F445DB1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF676C15D, CInt(&H596A), CInt(&H4CE2), &H82, &H34, &H33, &H99, &H6F, &H44, &H5D, &HB1)
 IID_IThumbnailCache = iid
End Function
Public Function IID_IThumbnailSettings() As UUID
'{F4376F00-BEF5-4d45-80F3-1E023BBF1209}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF4376F00, CInt(&HBEF5), CInt(&H4D45), &H80, &HF3, &H1E, &H2, &H3B, &HBF, &H12, &H9)
 IID_IThumbnailSettings = iid
End Function
Public Function IID_ITrackShellMenu() As UUID
'{8278F932-2A3E-11d2-838F-00C04FD918D0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8278F932, CInt(&H2A3E), CInt(&H11D2), &H83, &H8F, &H0, &HC0, &H4F, &HD9, &H18, &HD0)
 IID_ITrackShellMenu = iid
End Function

Public Function IID_IImageRecompress() As UUID
'{505f1513-6b3e-4892-a272-59f8889a4d3e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H505F1513, CInt(&H6B3E), CInt(&H4892), &HA2, &H72, &H59, &HF8, &H88, &H9A, &H4D, &H3E)
IID_IImageRecompress = iid
End Function
Public Function IID_ITranscodeImage() As UUID
'{BAE86DDD-DC11-421c-B7AB-CC55D1D65C44}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE86DDD, CInt(&HDC11), CInt(&H421C), &HB7, &HAB, &HCC, &H55, &HD1, &HD6, &H5C, &H44)
IID_ITranscodeImage = iid
End Function
Public Function IID_IParentAndItem() As UUID
'{b3a4b685-b685-4805-99d9-5dead2873236}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB3A4B685, CInt(&HB685), CInt(&H4805), &H99, &HD9, &H5D, &HEA, &HD2, &H87, &H32, &H36)
IID_IParentAndItem = iid
End Function
Public Function IID_ISearchBoxInfo() As UUID
'{6af6e03f-d664-4ef4-9626-f7e0ed36755e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6AF6E03F, CInt(&HD664), CInt(&H4EF4), &H96, &H26, &HF7, &HE0, &HED, &H36, &H75, &H5E)
IID_ISearchBoxInfo = iid
End Function
Public Function IID_IShellFolderViewCB() As UUID
'{2047E320-F2A9-11CE-AE65-08002B2E1262}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2047E320, CInt(&HF2A9), CInt(&H11CE), &HAE, &H65, &H8, &H0, &H2B, &H2E, &H12, &H62)
IID_IShellFolderViewCB = iid
End Function
Public Function IID_IPreviousVersionsInfo() As UUID
'{76e54780-ad74-48e3-a695-3ba9a0aff10d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H76E54780, CInt(&HAD74), CInt(&H48E3), &HA6, &H95, &H3B, &HA9, &HA0, &HAF, &HF1, &HD)
IID_IPreviousVersionsInfo = iid
End Function
Public Function IID_IZoneIdentifier() As UUID
'{cd45f185-1b21-48e2-967b-ead743a8914e}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCD45F185, CInt(&H1B21), CInt(&H48E2), &H96, &H7B, &HEA, &HD7, &H43, &HA8, &H91, &H4E)
IID_IZoneIdentifier = iid
End Function
Public Function IID_IApplicationAssociationRegistration() As UUID
'{4e530b0a-e611-4c77-a3ac-9031d022281b}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4E530B0A, CInt(&HE611), CInt(&H4C77), &HA3, &HAC, &H90, &H31, &HD0, &H22, &H28, &H1B)
IID_IApplicationAssociationRegistration = iid
End Function
Public Function IID_IApplicationAssociationRegistrationUI() As UUID
'{1f76a169-f994-40ac-8fc8-0959e8874710}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1F76A169, CInt(&HF994), CInt(&H40AC), &H8F, &HC8, &H9, &H59, &HE8, &H87, &H47, &H10)
IID_IApplicationAssociationRegistrationUI = iid
End Function
Public Function IID_ISystemInformation() As UUID
'{ADE87BF7-7B56-4275-8FAB-B9B0E591844B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HADE87BF7, CInt(&H7B56), CInt(&H4275), &H8F, &HAB, &HB9, &HB0, &HE5, &H91, &H84, &H4B)
IID_ISystemInformation = iid
End Function
Public Function IID_IFolderViewSettings() As UUID
'{ae8c987d-8797-4ed3-be72-2a47dd938db0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAE8C987D, CInt(&H8797), CInt(&H4ED3), &HBE, &H72, &H2A, &H47, &HDD, &H93, &H8D, &HB0)
IID_IFolderViewSettings = iid
End Function
Public Function IID_IFolderViewOptions() As UUID
'{3cc974d2-b302-4d36-ad3e-06d93f695d3f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CC974D2, CInt(&HB302), CInt(&H4D36), &HAD, &H3E, &H6, &HD9, &H3F, &H69, &H5D, &H3F)
IID_IFolderViewOptions = iid
End Function
Public Function IID_IResolveShellLink() As UUID
'{5cd52983-9449-11d2-963a-00c04f79adf0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CD52983, CInt(&H9449), CInt(&H11D2), &H96, &H3A, &H0, &HC0, &H4F, &H79, &HAD, &HF0)
IID_IResolveShellLink = iid
End Function
Public Function IID_IStartMenuPinnedList() As UUID
'{4CD19ADA-25A5-4A32-B3B7-347BEE5BE36B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4CD19ADA, CInt(&H25A5), CInt(&H4A32), &HB3, &HB7, &H34, &H7B, &HEE, &H5B, &HE3, &H6B)
IID_IStartMenuPinnedList = iid
End Function
Public Function IID_IObjMgr() As UUID
'{00BB2761-6A77-11D0-A535-00C04FD7D062}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBB2761, CInt(&H6A77), CInt(&H11D0), &HA5, &H35, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
IID_IObjMgr = iid
End Function
Public Function IID_IAutoCompleteDropDown() As UUID
'{3CD141F4-3C6A-11d2-BCAA-00C04FD929DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3CD141F4, CInt(&H3C6A), CInt(&H11D2), &HBC, &HAA, &H0, &HC0, &H4F, &HD9, &H29, &HDB)
IID_IAutoCompleteDropDown = iid
End Function
Public Function IID_IFolderFilter() As UUID
'{9CC22886-DC8E-11d2-B1D0-00C04F8EEB3E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9CC22886, CInt(&HDC8E), CInt(&H11D2), &HB1, &HD0, &H0, &HC0, &H4F, &H8E, &HEB, &H3E)
IID_IFolderFilter = iid
End Function
Public Function IID_IShellLinkDataList() As UUID
'{45e2b4ae-b1c3-11d0-b92f-00a0c90312e1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H45E2B4AE, CInt(&HB1C3), CInt(&H11D0), &HB9, &H2F, &H0, &HA0, &HC9, &H3, &H12, &HE1)
IID_IShellLinkDataList = iid
End Function
Public Function IID_IDataObjectAsyncCapability() As UUID
'{3D8B0590-F691-11d2-8EA9-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3D8B0590, CInt(&HF691), CInt(&H11D2), &H8E, &HA9, &H0, &H60, &H97, &HDF, &H5B, &HD4)
IID_IDataObjectAsyncCapability = iid
End Function
Public Function IID_IPortableDeviceManager() As UUID
'{A1567595-4C2F-4574-A6FA-ECEF917B9A40}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA1567595, CInt(&H4C2F), CInt(&H4574), &HA6, &HFA, &HEC, &HEF, &H91, &H7B, &H9A, &H40)
IID_IPortableDeviceManager = iid
End Function
Public Function IID_IPortableDeviceValuesCollection() As UUID
'{6E3F2D79-4E07-48C4-8208-D8C2E5AF4A99}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6E3F2D79, CInt(&H4E07), CInt(&H48C4), &H82, &H8, &HD8, &HC2, &HE5, &HAF, &H4A, &H99)
IID_IPortableDeviceValuesCollection = iid
End Function
Public Function IID_IPortableDevicePropVariantCollection() As UUID
'{89B2E422-4F1B-4316-BCEF-A44AFEA83EB3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89B2E422, CInt(&H4F1B), CInt(&H4316), &HBC, &HEF, &HA4, &H4A, &HFE, &HA8, &H3E, &HB3)
IID_IPortableDevicePropVariantCollection = iid
End Function
Public Function IID_IPortableDeviceKeyCollection() As UUID
'{DADA2357-E0AD-492E-98DB-DD61C53BA353}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDADA2357, CInt(&HE0AD), CInt(&H492E), &H98, &HDB, &HDD, &H61, &HC5, &H3B, &HA3, &H53)
IID_IPortableDeviceKeyCollection = iid
End Function
Public Function IID_IPortableDeviceValues() As UUID
'{6848F6F2-3155-4F86-B6F5-263EEEAB3143}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6848F6F2, CInt(&H3155), CInt(&H4F86), &HB6, &HF5, &H26, &H3E, &HEE, &HAB, &H31, &H43)
IID_IPortableDeviceValues = iid
End Function
Public Function IID_IPortableDevice() As UUID
'{625E2DF8-6392-4CF0-9AD1-3CFA5F17775C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H625E2DF8, CInt(&H6392), CInt(&H4CF0), &H9A, &HD1, &H3C, &HFA, &H5F, &H17, &H77, &H5C)
IID_IPortableDevice = iid
End Function
Public Function IID_IPortableDeviceContent() As UUID
'{6A96ED84-7C73-4480-9938-BF5AF477D426}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A96ED84, CInt(&H7C73), CInt(&H4480), &H99, &H38, &HBF, &H5A, &HF4, &H77, &HD4, &H26)
IID_IPortableDeviceContent = iid
End Function
Public Function IID_IEnumPortableDeviceObjectIDs() As UUID
'{10ECE955-CF41-4728-BFA0-41EEDF1BBF19}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10ECE955, CInt(&HCF41), CInt(&H4728), &HBF, &HA0, &H41, &HEE, &HDF, &H1B, &HBF, &H19)
IID_IEnumPortableDeviceObjectIDs = iid
End Function
Public Function IID_IPortableDeviceProperties() As UUID
'{7F6D695C-03DF-4439-A809-59266BEEE3A6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7F6D695C, CInt(&H3DF), CInt(&H4439), &HA8, &H9, &H59, &H26, &H6B, &HEE, &HE3, &HA6)
IID_IPortableDeviceProperties = iid
End Function
Public Function IID_IPortableDeviceResources() As UUID
'{FD8878AC-D841-4D17-891C-E6829CDB6934}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFD8878AC, CInt(&HD841), CInt(&H4D17), &H89, &H1C, &HE6, &H82, &H9C, &HDB, &H69, &H34)
IID_IPortableDeviceResources = iid
End Function
Public Function IID_IPortableDeviceCapabilities() As UUID
'{2C8C6DBF-E3DC-4061-BECC-8542E810D126}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C8C6DBF, CInt(&HE3DC), CInt(&H4061), &HBE, &HCC, &H85, &H42, &HE8, &H10, &HD1, &H26)
IID_IPortableDeviceCapabilities = iid
End Function
Public Function IID_IPortableDeviceService() As UUID
'{D3BD3A44-D7B5-40A9-98B7-2FA4D01DEC08}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD3BD3A44, CInt(&HD7B5), CInt(&H40A9), &H98, &HB7, &H2F, &HA4, &HD0, &H1D, &HEC, &H8)
IID_IPortableDeviceService = iid
End Function
Public Function IID_IPortableDeviceServiceCapabilities() As UUID
'{24DBD89D-413E-43E0-BD5B-197F3C56C886}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24DBD89D, CInt(&H413E), CInt(&H43E0), &HBD, &H5B, &H19, &H7F, &H3C, &H56, &HC8, &H86)
IID_IPortableDeviceServiceCapabilities = iid
End Function
Public Function IID_IPortableDeviceContent2() As UUID
'{9B4ADD96-F6BF-4034-8708-ECA72BF10554}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B4ADD96, CInt(&HF6BF), CInt(&H4034), &H87, &H8, &HEC, &HA7, &H2B, &HF1, &H5, &H54)
IID_IPortableDeviceContent2 = iid
End Function
Public Function IID_IPortableDeviceServiceMethods() As UUID
'{E20333C9-FD34-412D-A381-CC6F2D820DF7}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE20333C9, CInt(&HFD34), CInt(&H412D), &HA3, &H81, &HCC, &H6F, &H2D, &H82, &HD, &HF7)
IID_IPortableDeviceServiceMethods = iid
End Function
Public Function IID_IPortableDeviceDispatchFactory() As UUID
'{5E1EAFC3-E3D7-4132-96FA-759C0F9D1E0F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E1EAFC3, CInt(&HE3D7), CInt(&H4132), &H96, &HFA, &H75, &H9C, &HF, &H9D, &H1E, &HF)
IID_IPortableDeviceDispatchFactory = iid
End Function
Public Function IID_IWpdSerializer() As UUID
'{B32F4002-BB27-45FF-AF4F-06631C1E8DAD}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB32F4002, CInt(&HBB27), CInt(&H45FF), &HAF, &H4F, &H6, &H63, &H1C, &H1E, &H8D, &HAD)
IID_IWpdSerializer = iid
End Function
Public Function IID_IPortableDeviceDataStream() As UUID
'{88e04db3-1012-4d64-9996-f703a950d3f4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H88E04DB3, CInt(&H1012), CInt(&H4D64), &H99, &H96, &HF7, &H3, &HA9, &H50, &HD3, &HF4)
IID_IPortableDeviceDataStream = iid
End Function
Public Function IID_IPortableDeviceUnitsStream() As UUID
'{5e98025f-bfc4-47a2-9a5f-bc900a507c67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5E98025F, CInt(&HBFC4), CInt(&H47A2), &H9A, &H5F, &HBC, &H90, &HA, &H50, &H7C, &H67)
IID_IPortableDeviceUnitsStream = iid
End Function
Public Function IID_IPortableDevicePropertiesBulk() As UUID
'{482b05c0-4056-44ed-9e0f-5e23b009da93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H482B05C0, CInt(&H4056), CInt(&H44ED), &H9E, &HF, &H5E, &H23, &HB0, &H9, &HDA, &H93)
IID_IPortableDevicePropertiesBulk = iid
End Function
Public Function IID_IPortableDeviceServiceActivation() As UUID
'{e56b0534-d9b9-425c-9b99-75f97cb3d7c8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE56B0534, CInt(&HD9B9), CInt(&H425C), &H9B, &H99, &H75, &HF9, &H7C, &HB3, &HD7, &HC8)
IID_IPortableDeviceServiceActivation = iid
End Function
Public Function IID_IPortableDeviceWebControl() As UUID
'{94fc7953-5ca1-483a-8aee-df52e7747d00}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H94FC7953, CInt(&H5CA1), CInt(&H483A), &H8A, &HEE, &HDF, &H52, &HE7, &H74, &H7D, &H0)
IID_IPortableDeviceWebControl = iid
End Function
Public Function IID_IPortableDeviceServiceMethodCallback() As UUID
'{C424233C-AFCE-4828-A756-7ED7A2350083}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC424233C, CInt(&HAFCE), CInt(&H4828), &HA7, &H56, &H7E, &HD7, &HA2, &H35, &H0, &H83)
IID_IPortableDeviceServiceMethodCallback = iid
End Function
Public Function IID_IPortableDeviceServiceOpenCallback() As UUID
'{bced49c8-8efe-41ed-960b-61313abd47a9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCED49C8, CInt(&H8EFE), CInt(&H41ED), &H96, &HB, &H61, &H31, &H3A, &HBD, &H47, &HA9)
IID_IPortableDeviceServiceOpenCallback = iid
End Function
Public Function IID_IPortableDeviceEventCallback() As UUID
'{A8792A31-F385-493C-A893-40F64EB45F6E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA8792A31, CInt(&HF385), CInt(&H493C), &HA8, &H93, &H40, &HF6, &H4E, &HB4, &H5F, &H6E)
IID_IPortableDeviceEventCallback = iid
End Function
Public Function IID_IConnectionRequestCallback() As UUID
'{272C9AE0-7161-4AE0-91BD-9F448EE9C427}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H272C9AE0, CInt(&H7161), CInt(&H4AE0), &H91, &HBD, &H9F, &H44, &H8E, &HE9, &HC4, &H27)
IID_IConnectionRequestCallback = iid
End Function
Public Function IID_IPortableDevicePropertiesBulkCallback() As UUID
'{9deacb80-11e8-40e3-a9f3-f557986a7845}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9DEACB80, CInt(&H11E8), CInt(&H40E3), &HA9, &HF3, &HF5, &H57, &H98, &H6A, &H78, &H45)
IID_IPortableDevicePropertiesBulkCallback = iid
End Function
Public Function IID_IPortableDeviceConnector() As UUID
'{625E2DF8-6392-4CF0-9AD1-3CFA5F17775C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H625E2DF8, CInt(&H6392), CInt(&H4CF0), &H9A, &HD1, &H3C, &HFA, &H5F, &H17, &H77, &H5C)
IID_IPortableDeviceConnector = iid
End Function
Public Function IID_IEnumPortableDeviceConnectors() As UUID
'{BFDEF549-9247-454F-BD82-06FE80853FAA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBFDEF549, CInt(&H9247), CInt(&H454F), &HBD, &H82, &H6, &HFE, &H80, &H85, &H3F, &HAA)
IID_IEnumPortableDeviceConnectors = iid
End Function
Public Function IID_IEnumNetConnection() As UUID
'{C08956A0-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A0, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetConnection = iid
End Function
Public Function IID_INetConnection() As UUID
'{C08956A1-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A1, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnection = iid
End Function
Public Function IID_INetConnectionManager() As UUID
'{C08956A2-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A2, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnectionManager = iid
End Function
Public Function IID_INetConnectionConnectUi() As UUID
'{C08956A3-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956A3, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetConnectionConnectUi = iid
End Function
Public Function IID_IEnumNetSharingPortMapping() As UUID
'{C08956B0-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B0, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPortMapping = iid
End Function
Public Function IID_INetSharingPortMapping() As UUID
'{C08956B1-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B1, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingPortMapping = iid
End Function
Public Function IID_INetSharingPortMappingProps() As UUID
'{24B7E9B5-E38F-4685-851B-00892CF5F940}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24B7E9B5, CInt(&HE38F), CInt(&H4685), &H85, &H1B, &H0, &H89, &H2C, &HF5, &HF9, &H40)
IID_INetSharingPortMappingProps = iid
End Function
Public Function IID_IEnumNetSharingEveryConnection() As UUID
'{C08956B8-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B8, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingEveryConnection = iid
End Function
Public Function IID_IEnumNetSharingPublicConnection() As UUID
'{C08956B4-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B4, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPublicConnection = iid
End Function
Public Function IID_IEnumNetSharingPrivateConnection() As UUID
'{C08956B5-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B5, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_IEnumNetSharingPrivateConnection = iid
End Function
Public Function IID_INetSharingPortMappingCollection() As UUID
'{02E4A2DE-DA20-4E34-89C8-AC22275A010B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2E4A2DE, CInt(&HDA20), CInt(&H4E34), &H89, &HC8, &HAC, &H22, &H27, &H5A, &H1, &HB)
IID_INetSharingPortMappingCollection = iid
End Function
Public Function IID_INetConnectionProps() As UUID
'{F4277C95-CE5B-463D-8167-5662D9BCAA72}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF4277C95, CInt(&HCE5B), CInt(&H463D), &H81, &H67, &H56, &H62, &HD9, &HBC, &HAA, &H72)
IID_INetConnectionProps = iid
End Function
Public Function IID_INetSharingConfiguration() As UUID
'{C08956B6-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B6, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingConfiguration = iid
End Function
Public Function IID_INetSharingEveryConnectionCollection() As UUID
'{33C4643C-7811-46FA-A89A-768597BD7223}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H33C4643C, CInt(&H7811), CInt(&H46FA), &HA8, &H9A, &H76, &H85, &H97, &HBD, &H72, &H23)
IID_INetSharingEveryConnectionCollection = iid
End Function
Public Function IID_INetSharingPublicConnectionCollection() As UUID
'{7D7A6355-F372-4971-A149-BFC927BE762A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D7A6355, CInt(&HF372), CInt(&H4971), &HA1, &H49, &HBF, &HC9, &H27, &HBE, &H76, &H2A)
IID_INetSharingPublicConnectionCollection = iid
End Function
Public Function IID_INetSharingPrivateConnectionCollection() As UUID
'{38AE69E0-4409-402A-A2CB-E965C727F840}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H38AE69E0, CInt(&H4409), CInt(&H402A), &HA2, &HCB, &HE9, &H65, &HC7, &H27, &HF8, &H40)
IID_INetSharingPrivateConnectionCollection = iid
End Function
Public Function IID_INetSharingManager() As UUID
'{C08956B7-1CD3-11D1-B1C5-00805FC1270E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC08956B7, CInt(&H1CD3), CInt(&H11D1), &HB1, &HC5, &H0, &H80, &H5F, &HC1, &H27, &HE)
IID_INetSharingManager = iid
End Function
Public Function IID_IEnumReadyCallback() As UUID
'{61E00D45-8FFF-4e60-924E-6537B61612DD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H61E00D45, CInt(&H8FFF), CInt(&H4E60), &H92, &H4E, &H65, &H37, &HB6, &H16, &H12, &HDD)
 IID_IEnumReadyCallback = iid
End Function
Public Function IID_IEnumerableView() As UUID
'{8C8BF236-1AEC-495f-9894-91D57C3C686F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8C8BF236, CInt(&H1AEC), CInt(&H495F), &H98, &H94, &H91, &HD5, &H7C, &H3C, &H68, &H6F)
 IID_IEnumerableView = iid
End Function
Public Function IID_IPreviewItem() As UUID
'{36149969-0A8F-49c8-8B00-4AECB20222FB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H36149969, CInt(&HA8F), CInt(&H49C8), &H8B, &H0, &H4A, &HEC, &HB2, &H2, &H22, &HFB)
 IID_IPreviewItem = iid
End Function
Public Function IID_IViewStateIdentityItem() As UUID
'{9D264146-A94F-4195-9F9F-3BB12CE0C955}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9D264146, CInt(&HA94F), CInt(&H4195), &H9F, &H9F, &H3B, &HB1, &H2C, &HE0, &HC9, &H55)
 IID_IViewStateIdentityItem = iid
End Function
Public Function IID_IDisplayItem() As UUID
'{c6fd5997-9f6b-4888-8703-94e80e8cde3f}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6FD5997, CInt(&H9F6B), CInt(&H4888), &H87, &H3, &H94, &HE8, &HE, &H8C, &HDE, &H3F)
 IID_IDisplayItem = iid
End Function
Public Function IID_IUseToBrowseItem() As UUID
'{05edda5c-98a3-4717-8adb-c5e7da991eb1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5EDDA5C, CInt(&H98A3), CInt(&H4717), &H8A, &HDB, &HC5, &HE7, &HDA, &H99, &H1E, &HB1)
 IID_IUseToBrowseItem = iid
End Function
Public Function IID_ITransferMedium() As UUID
'{77f295d5-2d6f-4e19-b8ae-322f3e721ab5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H77F295D5, CInt(&H2D6F), CInt(&H4E19), &HB8, &HAE, &H32, &H2F, &H3E, &H72, &H1A, &HB5)
 IID_ITransferMedium = iid
End Function
Public Function IID_ICurrentItem() As UUID
'{240a7174-d653-4a1d-a6d3-d4943cfbfe3d}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H240A7174, CInt(&HD653), CInt(&H4A1D), &HA6, &HD3, &HD4, &H94, &H3C, &HFB, &HFE, &H3D)
 IID_ICurrentItem = iid
End Function
Public Function IID_IDelegateItem() As UUID
'{3c5a1c94-c951-4cb7-bb6d-3b93f30cce9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3C5A1C94, CInt(&HC951), CInt(&H4CB7), &HBB, &H6D, &H3B, &H93, &HF3, &HC, &HCE, &H9)
 IID_IDelegateItem = iid
End Function
Public Function IID_IIdentityName() As UUID
'{7d903fca-d6f9-4810-8332-946c0177e247}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D903FCA, CInt(&HD6F9), CInt(&H4810), &H83, &H32, &H94, &H6C, &H1, &H77, &HE2, &H47)
 IID_IIdentityName = iid
End Function
Public Function IID_IRelatedItem() As UUID
'{a73ce67a-8ab1-44f1-8d43-d2fcbf6b1cd0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA73CE67A, CInt(&H8AB1), CInt(&H44F1), &H8D, &H43, &HD2, &HFC, &HBF, &H6B, &H1C, &HD0)
 IID_IRelatedItem = iid
End Function
Public Function IID_IFilterCondition() As UUID
'{FCA2857D-1760-4AD3-8C63-C9B602FCBAEA}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFCA2857D, CInt(&H1760), CInt(&H4AD3), &H8C, &H63, &HC9, &HB6, &H2, &HFC, &HBA, &HEA)
 IID_IFilterCondition = iid
End Function
Public Function IID_IItemFilter() As UUID
'{7FCBEB25-ED60-45C9-9F5E-57B48493C4DD}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7FCBEB25, CInt(&HED60), CInt(&H45C9), &H9F, &H5E, &H57, &HB4, &H84, &H93, &HC4, &HDD)
 IID_IItemFilter = iid
End Function
Public Function IID_INewMenuClient() As UUID
'{dcb07fdc-3bb5-451c-90be-966644fed7b0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDCB07FDC, CInt(&H3BB5), CInt(&H451C), &H90, &HBE, &H96, &H66, &H44, &HFE, &HD7, &HB0)
 IID_INewMenuClient = iid
End Function
Public Function IID_IItemNameLimits() As UUID
'{1df0d7f1-b267-4d28-8b10-12e23202a5c4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1DF0D7F1, CInt(&HB267), CInt(&H4D28), &H8B, &H10, &H12, &HE2, &H32, &H2, &HA5, &HC4)
 IID_IItemNameLimits = iid
End Function
Public Function IID_ITaskFolderCollection() As UUID
'{79184A66-8664-423F-97F1-637356A5D812}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H79184A66, CInt(&H8664), CInt(&H423F), &H97, &HF1, &H63, &H73, &H56, &HA5, &HD8, &H12)
IID_ITaskFolderCollection = iid
End Function
Public Function IID_ITaskFolder() As UUID
'{8CFAC062-A080-4C15-9A88-AA7C2AF80DFC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CFAC062, CInt(&HA080), CInt(&H4C15), &H9A, &H88, &HAA, &H7C, &H2A, &HF8, &HD, &HFC)
IID_ITaskFolder = iid
End Function
Public Function IID_IRegisteredTask() As UUID
'{9C86F320-DEE3-4DD1-B972-A303F26B061E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9C86F320, CInt(&HDEE3), CInt(&H4DD1), &HB9, &H72, &HA3, &H3, &HF2, &H6B, &H6, &H1E)
IID_IRegisteredTask = iid
End Function
Public Function IID_IRunningTask() As UUID
'{653758FB-7B9A-4F1E-A471-BEEB8E9B834E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H653758FB, CInt(&H7B9A), CInt(&H4F1E), &HA4, &H71, &HBE, &HEB, &H8E, &H9B, &H83, &H4E)
IID_IRunningTask = iid
End Function
Public Function IID_IRunningTaskCollection() As UUID
'{6A67614B-6828-4FEC-AA54-6D52E8F1F2DB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6A67614B, CInt(&H6828), CInt(&H4FEC), &HAA, &H54, &H6D, &H52, &HE8, &HF1, &HF2, &HDB)
IID_IRunningTaskCollection = iid
End Function
Public Function IID_ITaskDefinition() As UUID
'{F5BC8FC5-536D-4F77-B852-FBC1356FDEB6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF5BC8FC5, CInt(&H536D), CInt(&H4F77), &HB8, &H52, &HFB, &HC1, &H35, &H6F, &HDE, &HB6)
IID_ITaskDefinition = iid
End Function
Public Function IID_IRegistrationInfo() As UUID
'{416D8B73-CB41-4EA1-805C-9BE9A5AC4A74}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H416D8B73, CInt(&HCB41), CInt(&H4EA1), &H80, &H5C, &H9B, &HE9, &HA5, &HAC, &H4A, &H74)
IID_IRegistrationInfo = iid
End Function
Public Function IID_ITriggerCollection() As UUID
'{85DF5081-1B24-4F32-878A-D9D14DF4CB77}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H85DF5081, CInt(&H1B24), CInt(&H4F32), &H87, &H8A, &HD9, &HD1, &H4D, &HF4, &HCB, &H77)
IID_ITriggerCollection = iid
End Function
Public Function IID_ITrigger() As UUID
'{09941815-EA89-4B5B-89E0-2A773801FAC3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9941815, CInt(&HEA89), CInt(&H4B5B), &H89, &HE0, &H2A, &H77, &H38, &H1, &HFA, &HC3)
IID_ITrigger = iid
End Function
Public Function IID_IRepetitionPattern() As UUID
'{7FB9ACF1-26BE-400E-85B5-294B9C75DFD6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7FB9ACF1, CInt(&H26BE), CInt(&H400E), &H85, &HB5, &H29, &H4B, &H9C, &H75, &HDF, &HD6)
IID_IRepetitionPattern = iid
End Function
Public Function IID_ITaskSettings() As UUID
'{8FD4711D-2D02-4C8C-87E3-EFF699DE127E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8FD4711D, CInt(&H2D02), CInt(&H4C8C), &H87, &HE3, &HEF, &HF6, &H99, &HDE, &H12, &H7E)
IID_ITaskSettings = iid
End Function
Public Function IID_IIdleSettings() As UUID
'{84594461-0053-4342-A8FD-088FABF11F32}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H84594461, CInt(&H53), CInt(&H4342), &HA8, &HFD, &H8, &H8F, &HAB, &HF1, &H1F, &H32)
IID_IIdleSettings = iid
End Function
Public Function IID_INetworkSettings() As UUID
'{9F7DEA84-C30B-4245-80B6-00E9F646F1B4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9F7DEA84, CInt(&HC30B), CInt(&H4245), &H80, &HB6, &H0, &HE9, &HF6, &H46, &HF1, &HB4)
IID_INetworkSettings = iid
End Function
Public Function IID_IPrincipal() As UUID
'{D98D51E5-C9B4-496A-A9C1-18980261CF0F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD98D51E5, CInt(&HC9B4), CInt(&H496A), &HA9, &HC1, &H18, &H98, &H2, &H61, &HCF, &HF)
IID_IPrincipal = iid
End Function
Public Function IID_IActionCollection() As UUID
'{02820E19-7B98-4ED2-B2E8-FDCCCEFF619B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2820E19, CInt(&H7B98), CInt(&H4ED2), &HB2, &HE8, &HFD, &HCC, &HCE, &HFF, &H61, &H9B)
IID_IActionCollection = iid
End Function
Public Function IID_IAction() As UUID
'{BAE54997-48B1-4CBE-9965-D6BE263EBEA4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBAE54997, CInt(&H48B1), CInt(&H4CBE), &H99, &H65, &HD6, &HBE, &H26, &H3E, &HBE, &HA4)
IID_IAction = iid
End Function
Public Function IID_IRegisteredTaskCollection() As UUID
'{86627EB4-42A7-41E4-A4D9-AC33A72F2D52}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H86627EB4, CInt(&H42A7), CInt(&H41E4), &HA4, &HD9, &HAC, &H33, &HA7, &H2F, &H2D, &H52)
IID_IRegisteredTaskCollection = iid
End Function
Public Function IID_ITaskService() As UUID
'{2FABA4C7-4DA9-4013-9697-20CC3FD40F85}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2FABA4C7, CInt(&H4DA9), CInt(&H4013), &H96, &H97, &H20, &HCC, &H3F, &HD4, &HF, &H85)
IID_ITaskService = iid
End Function
Public Function IID_ITaskHandler() As UUID
'{839D7762-5121-4009-9234-4F0D19394F04}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H839D7762, CInt(&H5121), CInt(&H4009), &H92, &H34, &H4F, &HD, &H19, &H39, &H4F, &H4)
IID_ITaskHandler = iid
End Function
Public Function IID_ITaskHandlerStatus() As UUID
'{EAEC7A8F-27A0-4DDC-8675-14726A01A38A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEAEC7A8F, CInt(&H27A0), CInt(&H4DDC), &H86, &H75, &H14, &H72, &H6A, &H1, &HA3, &H8A)
IID_ITaskHandlerStatus = iid
End Function
Public Function IID_ITaskVariables() As UUID
'{3E4C9351-D966-4B8B-BB87-CEBA68BB0107}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3E4C9351, CInt(&HD966), CInt(&H4B8B), &HBB, &H87, &HCE, &HBA, &H68, &HBB, &H1, &H7)
IID_ITaskVariables = iid
End Function
Public Function IID_ITaskNamedValuePair() As UUID
'{39038068-2B46-4AFD-8662-7BB6F868D221}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H39038068, CInt(&H2B46), CInt(&H4AFD), &H86, &H62, &H7B, &HB6, &HF8, &H68, &HD2, &H21)
IID_ITaskNamedValuePair = iid
End Function
Public Function IID_ITaskNamedValueCollection() As UUID
'{B4EF826B-63C3-46E4-A504-EF69E4F7EA4D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB4EF826B, CInt(&H63C3), CInt(&H46E4), &HA5, &H4, &HEF, &H69, &HE4, &HF7, &HEA, &H4D)
IID_ITaskNamedValueCollection = iid
End Function
Public Function IID_IIdleTrigger() As UUID
'{D537D2B0-9FB3-4D34-9739-1FF5CE7B1EF3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD537D2B0, CInt(&H9FB3), CInt(&H4D34), &H97, &H39, &H1F, &HF5, &HCE, &H7B, &H1E, &HF3)
IID_IIdleTrigger = iid
End Function
Public Function IID_ILogonTrigger() As UUID
'{72DADE38-FAE4-4B3E-BAF4-5D009AF02B1C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H72DADE38, CInt(&HFAE4), CInt(&H4B3E), &HBA, &HF4, &H5D, &H0, &H9A, &HF0, &H2B, &H1C)
IID_ILogonTrigger = iid
End Function
Public Function IID_ISessionStateChangeTrigger() As UUID
'{754DA71B-4385-4475-9DD9-598294FA3641}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H754DA71B, CInt(&H4385), CInt(&H4475), &H9D, &HD9, &H59, &H82, &H94, &HFA, &H36, &H41)
IID_ISessionStateChangeTrigger = iid
End Function
Public Function IID_IEventTrigger() As UUID
'{D45B0167-9653-4EEF-B94F-0732CA7AF251}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD45B0167, CInt(&H9653), CInt(&H4EEF), &HB9, &H4F, &H7, &H32, &HCA, &H7A, &HF2, &H51)
IID_IEventTrigger = iid
End Function
Public Function IID_ITimeTrigger() As UUID
'{B45747E0-EBA7-4276-9F29-85C5BB300006}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB45747E0, CInt(&HEBA7), CInt(&H4276), &H9F, &H29, &H85, &HC5, &HBB, &H30, &H0, &H6)
IID_ITimeTrigger = iid
End Function
Public Function IID_IDailyTrigger() As UUID
'{126C5CD8-B288-41D5-8DBF-E491446ADC5C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H126C5CD8, CInt(&HB288), CInt(&H41D5), &H8D, &HBF, &HE4, &H91, &H44, &H6A, &HDC, &H5C)
IID_IDailyTrigger = iid
End Function
Public Function IID_IWeeklyTrigger() As UUID
'{5038FC98-82FF-436D-8728-A512A57C9DC1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5038FC98, CInt(&H82FF), CInt(&H436D), &H87, &H28, &HA5, &H12, &HA5, &H7C, &H9D, &HC1)
IID_IWeeklyTrigger = iid
End Function
Public Function IID_IMonthlyTrigger() As UUID
'{97C45EF1-6B02-4A1A-9C0E-1EBFBA1500AC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H97C45EF1, CInt(&H6B02), CInt(&H4A1A), &H9C, &HE, &H1E, &HBF, &HBA, &H15, &H0, &HAC)
IID_IMonthlyTrigger = iid
End Function
Public Function IID_IMonthlyDOWTrigger() As UUID
'{77D025A3-90FA-43AA-B52E-CDA5499B946A}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H77D025A3, CInt(&H90FA), CInt(&H43AA), &HB5, &H2E, &HCD, &HA5, &H49, &H9B, &H94, &H6A)
IID_IMonthlyDOWTrigger = iid
End Function
Public Function IID_IBootTrigger() As UUID
'{2A9C35DA-D357-41F4-BBC1-207AC1B1F3CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2A9C35DA, CInt(&HD357), CInt(&H41F4), &HBB, &HC1, &H20, &H7A, &HC1, &HB1, &HF3, &HCB)
IID_IBootTrigger = iid
End Function
Public Function IID_IRegistrationTrigger() As UUID
'{4C8FEC3A-C218-4E0C-B23D-629024DB91A2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C8FEC3A, CInt(&HC218), CInt(&H4E0C), &HB2, &H3D, &H62, &H90, &H24, &HDB, &H91, &HA2)
IID_IRegistrationTrigger = iid
End Function
Public Function IID_IExecAction() As UUID
'{4C3D624D-FD6B-49A3-B9B7-09CB3CD3F047}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4C3D624D, CInt(&HFD6B), CInt(&H49A3), &HB9, &HB7, &H9, &HCB, &H3C, &HD3, &HF0, &H47)
IID_IExecAction = iid
End Function
Public Function IID_IExecAction2() As UUID
'{F2A82542-BDA5-4E6B-9143-E2BF4F8987B6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF2A82542, CInt(&HBDA5), CInt(&H4E6B), &H91, &H43, &HE2, &HBF, &H4F, &H89, &H87, &HB6)
IID_IExecAction2 = iid
End Function
Public Function IID_IShowMessageAction() As UUID
'{505E9E68-AF89-46B8-A30F-56162A83D537}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H505E9E68, CInt(&HAF89), CInt(&H46B8), &HA3, &HF, &H56, &H16, &H2A, &H83, &HD5, &H37)
IID_IShowMessageAction = iid
End Function
Public Function IID_IComHandlerAction() As UUID
'{6D2FD252-75C5-4F66-90BA-2A7D8CC3039F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6D2FD252, CInt(&H75C5), CInt(&H4F66), &H90, &HBA, &H2A, &H7D, &H8C, &HC3, &H3, &H9F)
IID_IComHandlerAction = iid
End Function
Public Function IID_IEmailAction() As UUID
'{10F62C64-7E16-4314-A0C2-0C3683F99D40}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H10F62C64, CInt(&H7E16), CInt(&H4314), &HA0, &HC2, &HC, &H36, &H83, &HF9, &H9D, &H40)
IID_IEmailAction = iid
End Function
Public Function IID_IPrincipal2() As UUID
'{248919AE-E345-4A6D-8AEB-E0D3165C904E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H248919AE, CInt(&HE345), CInt(&H4A6D), &H8A, &HEB, &HE0, &HD3, &H16, &H5C, &H90, &H4E)
IID_IPrincipal2 = iid
End Function
Public Function IID_ITaskSettings2() As UUID
'{2C05C3F0-6EED-4C05-A15F-ED7D7A98A369}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2C05C3F0, CInt(&H6EED), CInt(&H4C05), &HA1, &H5F, &HED, &H7D, &H7A, &H98, &HA3, &H69)
IID_ITaskSettings2 = iid
End Function
Public Function IID_ITaskSettings3() As UUID
'{0AD9D0D7-0C7F-4EBB-9A5F-D1C648DCA528}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAD9D0D7, CInt(&HC7F), CInt(&H4EBB), &H9A, &H5F, &HD1, &HC6, &H48, &HDC, &HA5, &H28)
IID_ITaskSettings3 = iid
End Function
Public Function IID_IMaintenanceSettings() As UUID
'{A6024FA8-9652-4ADB-A6BF-5CFCD877A7BA}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA6024FA8, CInt(&H9652), CInt(&H4ADB), &HA6, &HBF, &H5C, &HFC, &HD8, &H77, &HA7, &HBA)
IID_IMaintenanceSettings = iid
End Function

Public Function IID_IStream() As UUID
'{0000000C-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IStream = iid
End Function

Public Function IID_IUnknown() As UUID
'"{00000000-0000-0000-C000-000000000046}"
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H0, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
  IID_IUnknown = iid

End Function



Public Function BHID_AssociationArray() As UUID
'DEFINE_GUID(BHID_AssociationArray, 0xBEA9EF17, 0x82F1, 0x4F60, 0x92,0x84, 0x4F,0x8D,0xB7,0x5C,0x3B,0xE9)
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBEA9EF17, &H82F1, &H4F60, &H92, &H84, &H4F, &H8D, &HB7, &H5C, &H3B, &HE9)
  BHID_AssociationArray = iid
End Function

Public Function BHID_SFUIObject() As UUID
'DEFINE_GUID(BHID_SFUIObject,  0x3981E225, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5);
'{3981e225-f559-11d3-8e3a-00c04f6837d5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E225, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
  BHID_SFUIObject = iid
End Function
Public Function BHID_DataObject() As UUID
'{0xB8C0BD9F, 0xED24, 0x455C, 0x83,0xE6, 0xD5,0x39,0x0C,0x4F,0xE8,0xC4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8C0BD9F, &HED24, &H455C, &H83, &HE6, &HD5, &H39, &HC, &H4F, &HE8, &HC4)
 BHID_DataObject = iid
End Function
Public Function BHID_SFObject() As UUID
'{0x3981E224, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E224, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_SFObject = iid
End Function
Public Function BHID_SFViewObject() As UUID
'{0x3981E226, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E226, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_SFViewObject = iid
End Function
Public Function BHID_Storage() As UUID
'{0x3981E227, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E227, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_Storage = iid
End Function
Public Function BHID_Stream() As UUID
'{0x1CEBB3AB, 0x7C10, 0x499A, 0xA4,0x17, 0x92,0xCA,0x16,0xC4,0xCB,0x83}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1CEBB3AB, &H7C10, &H499A, &HA4, &H17, &H92, &HCA, &H16, &HC4, &HCB, &H83)
 BHID_Stream = iid
End Function
Public Function BHID_StorageEnum() As UUID
'{0x4621A4E3, 0xF0D6, 0x4773, 0x8A,0x9C, 0x46,0xE7,0x7B,0x17,0x48,0x40}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4621A4E3, &HF0D6, &H4773, &H8A, &H9C, &H46, &HE7, &H7B, &H17, &H48, &H40)
 BHID_StorageEnum = iid
End Function
Public Function BHID_Transfer() As UUID
'{0xD5E346A1, 0xF753, 0x4932, 0xB4,0x03, 0x45,0x74,0x80,0x0E,0x24,0x98}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD5E346A1, &HF753, &H4932, &HB4, &H3, &H45, &H74, &H80, &HE, &H24, &H98)
 BHID_Transfer = iid
End Function
Public Function BHID_Filter() As UUID
'{0x38D08778, 0xF557, 0x4690, 0x9E,0xBF, 0xBA,0x54,0x70,0x6A,0xD8,0xF7}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H38D08778, &HF557, &H4690, &H9E, &HBF, &HBA, &H54, &H70, &H6A, &HD8, &HF7)
 BHID_Filter = iid
End Function
Public Function BHID_LinkTargetItem() As UUID
'{0x3981E228, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E228, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
 BHID_LinkTargetItem = iid
End Function
Public Function BHID_PropertyStore() As UUID
'{0x0384E1A4, 0x1523, 0x439C, 0xA4,0xC8, 0xAB,0x91,0x10,0x52,0xF5,0x86}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H384E1A4, &H1523, &H439C, &HA4, &HC8, &HAB, &H91, &H10, &H52, &HF5, &H86)
 BHID_PropertyStore = iid
End Function
Public Function BHID_EnumAssocHandlers() As UUID
'{0xB8AB0B9C, 0xC2EC, 0x4F7A, 0x91,0x8D, 0x31,0x49,0x00,0xE6,0x28,0x0A}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8AB0B9C, &HC2EC, &H4F7A, &H91, &H8D, &H31, &H49, &H0, &HE6, &H28, &HA)
 BHID_EnumAssocHandlers = iid
End Function
Public Function BHID_ThumbnailHandler() As UUID
'{0x7B2E650A, 0x8E20, 0x4F4A, 0xB0,0x9E, 0x65,0x97,0xAF,0xC7,0x2F,0xB0}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B2E650A, &H8E20, &H4F4A, &HB0, &H9E, &H65, &H97, &HAF, &HC7, &H2F, &HB0)
 BHID_ThumbnailHandler = iid
End Function
Public Function BHID_EnumItems() As UUID
'{0x94F60519, 0x2850, 0x4924, 0xAA,0x5A, 0xD1,0x5E,0x84,0x86,0x80,0x39}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94F60519, &H2850, &H4924, &HAA, &H5A, &HD1, &H5E, &H84, &H86, &H80, &H39)
 BHID_EnumItems = iid
End Function
Public Function BHID_RandomAccessStream() As UUID
'0xf16fc93b, 0x77ae, 0x4cfe, 0xbd, 0xa7, 0xa8, 0x66, 0xee, 0xa6, 0x87, 0x8d
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF16FC93B, &H77AE, &H4CFE, &HBD, &HA7, &HA8, &H66, &HEE, &HA6, &H87, &H8D)
 BHID_RandomAccessStream = iid
End Function
Public Function BHID_FilePlaceholder() As UUID
'0x8677dceb, 0xaae0, 0x4005, 0x8d, 0x3d, 0x54, 0x7f, 0xa8, 0x52, 0xf8, 0x25)
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8677DCEB, &HAAE0, &H4005, &H8D, &H3D, &H54, &H7F, &HA8, &H52, &HF8, &H25)
 BHID_FilePlaceholder = iid
End Function
Public Function IID_IShellIconOverlay() As UUID
'{7d688a70-c613-11d0-999b-00c04fd655e1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D688A70, CInt(&HC613), CInt(&H11D0), &H99, &H9B, &H0, &HC0, &H4F, &HD6, &H55, &HE1)
 IID_IShellIconOverlay = iid
End Function
Public Function IID_IShellIconOverlayIdentifier() As UUID
'{0c6c4200-c589-11d0-999a-00c04fd655e1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC6C4200, CInt(&HC589), CInt(&H11D0), &H99, &H9A, &H0, &HC0, &H4F, &HD6, &H55, &HE1)
 IID_IShellIconOverlayIdentifier = iid
End Function
Public Function IID_IListView() As UUID
'{E5B16AF2-3990-4681-A609-1F060CD14269}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE5B16AF2, CInt(&H3990), CInt(&H4681), &HA6, &H9, &H1F, &H6, &HC, &HD1, &H42, &H69)
 IID_IListView = iid
End Function
Public Function IID_IListViewFooter() As UUID
'{F0034DA8-8A22-4151-8F16-2EBA76565BCC}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF0034DA8, CInt(&H8A22), CInt(&H4151), &H8F, &H16, &H2E, &HBA, &H76, &H56, &H5B, &HCC)
 IID_IListViewFooter = iid
End Function
Public Function IID_IListViewFooterCallback() As UUID
'{88EB9442-913B-4AB4-A741-DD99DCB7558B}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H88EB9442, CInt(&H913B), CInt(&H4AB4), &HA7, &H41, &HDD, &H99, &HDC, &HB7, &H55, &H8B)
 IID_IListViewFooterCallback = iid
End Function
Public Function IID_IOwnerDataCallback() As UUID
'{44C09D56-8D3B-419D-A462-7B956B105B47}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H44C09D56, CInt(&H8D3B), CInt(&H419D), &HA4, &H62, &H7B, &H95, &H6B, &H10, &H5B, &H47)
 IID_IOwnerDataCallback = iid
End Function
Public Function IID_IPropertyControlBase() As UUID
'{6E71A510-732A-4557-9596-A827E36DAF8F}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6E71A510, CInt(&H732A), CInt(&H4557), &H95, &H96, &HA8, &H27, &HE3, &H6D, &HAF, &H8F)
 IID_IPropertyControlBase = iid
End Function
Public Function IID_IPropertyControl() As UUID
'{5E82A4DD-9561-476A-8634-1BEBACBA4A38}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E82A4DD, CInt(&H9561), CInt(&H476A), &H86, &H34, &H1B, &HEB, &HAC, &HBA, &H4A, &H38)
 IID_IPropertyControl = iid
End Function
Public Function IID_IDrawPropertyControl() As UUID
'{E6DFF6FD-BCD5-4162-9C65-A3B18C616FDB}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HE6DFF6FD, CInt(&HBCD5), CInt(&H4162), &H9C, &H65, &HA3, &HB1, &H8C, &H61, &H6F, &HDB)
 IID_IDrawPropertyControl = iid
End Function
Public Function IID_IPropertyValue() As UUID
'{7AF7F355-1066-4E17-B1F2-19FE2F099CD2}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7AF7F355, CInt(&H1066), CInt(&H4E17), &HB1, &HF2, &H19, &HFE, &H2F, &H9, &H9C, &HD2)
 IID_IPropertyValue = iid
End Function
Public Function IID_ISubItemCallback() As UUID
'{11A66240-5489-42C2-AEBF-286FC831524C}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H11A66240, CInt(&H5489), CInt(&H42C2), &HAE, &HBF, &H28, &H6F, &HC8, &H31, &H52, &H4C)
 IID_ISubItemCallback = iid
End Function

Public Function IID_IShellApp() As UUID
'{A3E14960-935F-11D1-B8B8-006008059382}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA3E14960, CInt(&H935F), CInt(&H11D1), &HB8, &HB8, &H0, &H60, &H8, &H5, &H93, &H82)
 IID_IShellApp = iid
End Function
Public Function IID_IAppPublisher() As UUID
'{07250A10-9CF9-11D1-9076-006008059382}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7250A10, CInt(&H9CF9), CInt(&H11D1), &H90, &H76, &H0, &H60, &H8, &H5, &H93, &H82)
 IID_IAppPublisher = iid
End Function
Public Function IID_IBandSite() As UUID
'{4CF504B0-DE96-11D0-8B3F-00A0C911E8E5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CF504B0, CInt(&HDE96), CInt(&H11D0), &H8B, &H3F, &H0, &HA0, &HC9, &H11, &HE8, &HE5)
 IID_IBandSite = iid
End Function
Public Function IID_INewWindowManager() As UUID
'{4CF504B0-DE96-11D0-8B3F-00A0C911E8E5}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4CF504B0, CInt(&HDE96), CInt(&H11D0), &H8B, &H3F, &H0, &HA0, &HC9, &H11, &HE8, &HE5)
 IID_INewWindowManager = iid
End Function
Public Function IID_IDelegateFolder() As UUID
'{ADD8BA80-002B-11D0-8F0F-00C04FD7D062}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HADD8BA80, CInt(&H2B), CInt(&H11D0), &H8F, &HF, &H0, &HC0, &H4F, &HD7, &HD0, &H62)
 IID_IDelegateFolder = iid
End Function
Public Function IID_IBrowserFrameOptions() As UUID
'{10DF43C8-1DBE-11d3-8B34-006097DF5BD4}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10DF43C8, CInt(&H1DBE), CInt(&H11D3), &H8B, &H34, &H0, &H60, &H97, &HDF, &H5B, &HD4)
 IID_IBrowserFrameOptions = iid
End Function
Public Function IID_IFileIsInUse() As UUID
'{64a1cbf0-3a1a-4461-9158-376969693950}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H64A1CBF0, CInt(&H3A1A), CInt(&H4461), &H91, &H58, &H37, &H69, &H69, &H69, &H39, &H50)
 IID_IFileIsInUse = iid
End Function
Public Function IID_IOpenControlPanel() As UUID
'{D11AD862-66DE-4DF4-BF6C-1F5621996AF1}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD11AD862, CInt(&H66DE), CInt(&H4DF4), &HBF, &H6C, &H1F, &H56, &H21, &H99, &H6A, &HF1)
 IID_IOpenControlPanel = iid
End Function

Public Function SID_STopLevelBrowser() As UUID
'{4C96BE40-915C-11CF-99D3-00AA004AE837}

Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C96BE40, CInt(&H915C), CInt(&H11CF), &H99, &HD3, &H0, &HAA, &H0, &H4A, &HE8, &H37)
 SID_STopLevelBrowser = iid
End Function
Public Function SID_SExplorerBrowserFrame() As UUID
SID_SExplorerBrowserFrame = IID_ICommDlgBrowser
End Function
Public Function SID_SFolderView() As UUID
SID_SFolderView = IID_IFolderView
End Function
Public Function SID_SProfferService() As UUID
SID_SProfferService = IID_IProfferService
End Function
Public Function SID_WizardHost() As UUID
SID_WizardHost = IID_IWebWizardExtension
End Function
Public Function SID_CDWizardHost() As UUID
SID_CDWizardHost = IID_ICDBurnExt
End Function
Public Function SID_SBandSite() As UUID
SID_SBandSite = IID_IBandSite
End Function
Public Function SID_SNewMenuClient() As UUID
SID_SNewMenuClient = IID_INewMenuClient
End Function
Public Function SID_SNewWindowManager() As UUID
SID_SNewWindowManager = IID_INewWindowManager
End Function
Public Function SID_ExecuteCommandHost() As UUID
SID_ExecuteCommandHost = IID_IExecuteCommandHost
End Function
Public Function FOLDERID_NetworkFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD20BEEC4, CInt(&H5CA8), CInt(&H4905), &HAE, &H3B, &HBF, &H25, &H1E, &HA0, &H9B, &H53)
 FOLDERID_NetworkFolder = iid
End Function

Public Function FOLDERID_ComputerFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC0837C, CInt(&HBBF8), CInt(&H452A), &H85, &HD, &H79, &HD0, &H8E, &H66, &H7C, &HA7)
 FOLDERID_ComputerFolder = iid
End Function

Public Function FOLDERID_InternetFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4D9F7874, CInt(&H4E0C), CInt(&H4904), &H96, &H7B, &H40, &HB0, &HD2, &HC, &H3E, &H4B)
 FOLDERID_InternetFolder = iid
End Function

Public Function FOLDERID_ControlPanelFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82A74AEB, CInt(&HAEB4), CInt(&H465C), &HA0, &H14, &HD0, &H97, &HEE, &H34, &H6D, &H63)
 FOLDERID_ControlPanelFolder = iid
End Function

Public Function FOLDERID_PrintersFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H76FC4E2D, CInt(&HD6AD), CInt(&H4519), &HA6, &H63, &H37, &HBD, &H56, &H6, &H81, &H85)
 FOLDERID_PrintersFolder = iid
End Function

Public Function FOLDERID_SyncManagerFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43668BF8, CInt(&HC14E), CInt(&H49B2), &H97, &HC9, &H74, &H77, &H84, &HD7, &H84, &HB7)
 FOLDERID_SyncManagerFolder = iid
End Function

Public Function FOLDERID_SyncSetupFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF214138, CInt(&HB1D3), CInt(&H4A90), &HBB, &HA9, &H27, &HCB, &HC0, &HC5, &H38, &H9A)
 FOLDERID_SyncSetupFolder = iid
End Function

Public Function FOLDERID_ConflictFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4BFEFB45, CInt(&H347D), CInt(&H4006), &HA5, &HBE, &HAC, &HC, &HB0, &H56, &H71, &H92)
 FOLDERID_ConflictFolder = iid
End Function

Public Function FOLDERID_SyncResultsFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H289A9A43, CInt(&HBE44), CInt(&H4057), &HA4, &H1B, &H58, &H7A, &H76, &HD7, &HE7, &HF9)
 FOLDERID_SyncResultsFolder = iid
End Function

Public Function FOLDERID_RecycleBinFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7534046, CInt(&H3ECB), CInt(&H4C18), &HBE, &H4E, &H64, &HCD, &H4C, &HB7, &HD6, &HAC)
 FOLDERID_RecycleBinFolder = iid
End Function

Public Function FOLDERID_ConnectionsFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F0CD92B, CInt(&H2E97), CInt(&H45D1), &H88, &HFF, &HB0, &HD1, &H86, &HB8, &HDE, &HDD)
 FOLDERID_ConnectionsFolder = iid
End Function

Public Function FOLDERID_Fonts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFD228CB7, CInt(&HAE11), CInt(&H4AE3), &H86, &H4C, &H16, &HF3, &H91, &HA, &HB8, &HFE)
 FOLDERID_Fonts = iid
End Function

Public Function FOLDERID_Desktop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4BFCC3A, CInt(&HDB2C), CInt(&H424C), &HB0, &H29, &H7F, &HE9, &H9A, &H87, &HC6, &H41)
 FOLDERID_Desktop = iid
End Function

Public Function FOLDERID_Startup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB97D20BB, CInt(&HF46A), CInt(&H4C97), &HBA, &H10, &H5E, &H36, &H8, &H43, &H8, &H54)
 FOLDERID_Startup = iid
End Function

Public Function FOLDERID_Programs() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA77F5D77, CInt(&H2E2B), CInt(&H44C3), &HA6, &HA2, &HAB, &HA6, &H1, &H5, &H4A, &H51)
 FOLDERID_Programs = iid
End Function

Public Function FOLDERID_StartMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H625B53C3, CInt(&HAB48), CInt(&H4EC1), &HBA, &H1F, &HA1, &HEF, &H41, &H46, &HFC, &H19)
 FOLDERID_StartMenu = iid
End Function

Public Function FOLDERID_Recent() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAE50C081, CInt(&HEBD2), CInt(&H438A), &H86, &H55, &H8A, &H9, &H2E, &H34, &H98, &H7A)
 FOLDERID_Recent = iid
End Function

Public Function FOLDERID_SendTo() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8983036C, CInt(&H27C0), CInt(&H404B), &H8F, &H8, &H10, &H2D, &H10, &HDC, &HFD, &H74)
 FOLDERID_SendTo = iid
End Function

Public Function FOLDERID_Documents() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HFDD39AD0, CInt(&H238F), CInt(&H46AF), &HAD, &HB4, &H6C, &H85, &H48, &H3, &H69, &HC7)
 FOLDERID_Documents = iid
End Function

Public Function FOLDERID_Favorites() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1777F761, CInt(&H68AD), CInt(&H4D8A), &H87, &HBD, &H30, &HB7, &H59, &HFA, &H33, &HDD)
 FOLDERID_Favorites = iid
End Function

Public Function FOLDERID_NetHood() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC5ABBF53, CInt(&HE17F), CInt(&H4121), &H89, &H0, &H86, &H62, &H6F, &HC2, &HC9, &H73)
 FOLDERID_NetHood = iid
End Function

Public Function FOLDERID_PrintHood() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9274BD8D, CInt(&HCFD1), CInt(&H41C3), &HB3, &H5E, &HB1, &H3F, &H55, &HA7, &H58, &HF4)
 FOLDERID_PrintHood = iid
End Function

Public Function FOLDERID_Templates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA63293E8, CInt(&H664E), CInt(&H48DB), &HA0, &H79, &HDF, &H75, &H9E, &H5, &H9, &HF7)
 FOLDERID_Templates = iid
End Function

Public Function FOLDERID_CommonStartup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H82A5EA35, CInt(&HD9CD), CInt(&H47C5), &H96, &H29, &HE1, &H5D, &H2F, &H71, &H4E, &H6E)
 FOLDERID_CommonStartup = iid
End Function

Public Function FOLDERID_CommonPrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H139D44E, CInt(&H6AFE), CInt(&H49F2), &H86, &H90, &H3D, &HAF, &HCA, &HE6, &HFF, &HB8)
 FOLDERID_CommonPrograms = iid
End Function

Public Function FOLDERID_CommonStartMenu() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA4115719, CInt(&HD62E), CInt(&H491D), &HAA, &H7C, &HE7, &H4B, &H8B, &HE3, &HB0, &H67)
 FOLDERID_CommonStartMenu = iid
End Function

Public Function FOLDERID_PublicDesktop() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4AA340D, CInt(&HF20F), CInt(&H4863), &HAF, &HEF, &HF8, &H7E, &HF2, &HE6, &HBA, &H25)
 FOLDERID_PublicDesktop = iid
End Function

Public Function FOLDERID_ProgramData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H62AB5D82, CInt(&HFDC1), CInt(&H4DC3), &HA9, &HDD, &H7, &HD, &H1D, &H49, &H5D, &H97)
 FOLDERID_ProgramData = iid
End Function

Public Function FOLDERID_CommonTemplates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB94237E7, CInt(&H57AC), CInt(&H4347), &H91, &H51, &HB0, &H8C, &H6C, &H32, &HD1, &HF7)
 FOLDERID_CommonTemplates = iid
End Function

Public Function FOLDERID_PublicDocuments() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HED4824AF, CInt(&HDCE4), CInt(&H45A8), &H81, &HE2, &HFC, &H79, &H65, &H8, &H36, &H34)
 FOLDERID_PublicDocuments = iid
End Function

Public Function FOLDERID_RoamingAppData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3EB685DB, CInt(&H65F9), CInt(&H4CF6), &HA0, &H3A, &HE3, &HEF, &H65, &H72, &H9F, &H3D)
 FOLDERID_RoamingAppData = iid
End Function

Public Function FOLDERID_LocalAppData() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF1B32785, CInt(&H6FBA), CInt(&H4FCF), &H9D, &H55, &H7B, &H8E, &H7F, &H15, &H70, &H91)
 FOLDERID_LocalAppData = iid
End Function

Public Function FOLDERID_LocalAppDataLow() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA520A1A4, CInt(&H1780), CInt(&H4FF6), &HBD, &H18, &H16, &H73, &H43, &HC5, &HAF, &H16)
 FOLDERID_LocalAppDataLow = iid
End Function

Public Function FOLDERID_InternetCache() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H352481E8, CInt(&H33BE), CInt(&H4251), &HBA, &H85, &H60, &H7, &HCA, &HED, &HCF, &H9D)
 FOLDERID_InternetCache = iid
End Function

Public Function FOLDERID_Cookies() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2B0F765D, CInt(&HC0E9), CInt(&H4171), &H90, &H8E, &H8, &HA6, &H11, &HB8, &H4F, &HF6)
 FOLDERID_Cookies = iid
End Function

Public Function FOLDERID_History() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD9DC8A3B, CInt(&HB784), CInt(&H432E), &HA7, &H81, &H5A, &H11, &H30, &HA7, &H59, &H63)
 FOLDERID_History = iid
End Function

Public Function FOLDERID_System() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AC14E77, CInt(&H2E7), CInt(&H4E5D), &HB7, &H44, &H2E, &HB1, &HAE, &H51, &H98, &HB7)
 FOLDERID_System = iid
End Function

Public Function FOLDERID_SystemX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD65231B0, CInt(&HB2F1), CInt(&H4857), &HA4, &HCE, &HA8, &HE7, &HC6, &HEA, &H7D, &H27)
 FOLDERID_SystemX86 = iid
End Function

Public Function FOLDERID_Windows() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF38BF404, CInt(&H1D43), CInt(&H42F2), &H93, &H5, &H67, &HDE, &HB, &H28, &HFC, &H23)
 FOLDERID_Windows = iid
End Function

Public Function FOLDERID_Profile() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E6C858F, CInt(&HE22), CInt(&H4760), &H9A, &HFE, &HEA, &H33, &H17, &HB6, &H71, &H73)
 FOLDERID_Profile = iid
End Function

Public Function FOLDERID_Pictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H33E28130, CInt(&H4E1E), CInt(&H4676), &H83, &H5A, &H98, &H39, &H5C, &H3B, &HC3, &HBB)
 FOLDERID_Pictures = iid
End Function

Public Function FOLDERID_ProgramFilesX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7C5A40EF, CInt(&HA0FB), CInt(&H4BFC), &H87, &H4A, &HC0, &HF2, &HE0, &HB9, &HFA, &H8E)
 FOLDERID_ProgramFilesX86 = iid
End Function

Public Function FOLDERID_ProgramFilesCommonX86() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE974D24, CInt(&HD9C6), CInt(&H4D3E), &HBF, &H91, &HF4, &H45, &H51, &H20, &HB9, &H17)
 FOLDERID_ProgramFilesCommonX86 = iid
End Function

Public Function FOLDERID_ProgramFilesX64() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6D809377, CInt(&H6AF0), CInt(&H444B), &H89, &H57, &HA3, &H77, &H3F, &H2, &H20, &HE)
 FOLDERID_ProgramFilesX64 = iid
End Function

Public Function FOLDERID_ProgramFilesCommonX64() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6365D5A7, CInt(&HF0D), CInt(&H45E5), &H87, &HF6, &HD, &HA5, &H6B, &H6A, &H4F, &H7D)
 FOLDERID_ProgramFilesCommonX64 = iid
End Function

Public Function FOLDERID_ProgramFiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H905E63B6, CInt(&HC1BF), CInt(&H494E), &HB2, &H9C, &H65, &HB7, &H32, &HD3, &HD2, &H1A)
 FOLDERID_ProgramFiles = iid
End Function

Public Function FOLDERID_ProgramFilesCommon() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF7F1ED05, CInt(&H9F6D), CInt(&H47A2), &HAA, &HAE, &H29, &HD3, &H17, &HC6, &HF0, &H66)
 FOLDERID_ProgramFilesCommon = iid
End Function

Public Function FOLDERID_AdminTools() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H724EF170, CInt(&HA42D), CInt(&H4FEF), &H9F, &H26, &HB6, &HE, &H84, &H6F, &HBA, &H4F)
 FOLDERID_AdminTools = iid
End Function

Public Function FOLDERID_CommonAdminTools() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HD0384E7D, CInt(&HBAC3), CInt(&H4797), &H8F, &H14, &HCB, &HA2, &H29, &HB3, &H92, &HB5)
 FOLDERID_CommonAdminTools = iid
End Function

Public Function FOLDERID_Music() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4BD8D571, CInt(&H6D19), CInt(&H48D3), &HBE, &H97, &H42, &H22, &H20, &H8, &HE, &H43)
 FOLDERID_Music = iid
End Function

Public Function FOLDERID_Videos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H18989B1D, CInt(&H99B5), CInt(&H455B), &H84, &H1C, &HAB, &H7C, &H74, &HE4, &HDD, &HFC)
 FOLDERID_Videos = iid
End Function

Public Function FOLDERID_PublicPictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB6EBFB86, CInt(&H6907), CInt(&H413C), &H9A, &HF7, &H4F, &HC2, &HAB, &HF0, &H7C, &HC5)
 FOLDERID_PublicPictures = iid
End Function

Public Function FOLDERID_PublicMusic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3214FAB5, CInt(&H9757), CInt(&H4298), &HBB, &H61, &H92, &HA9, &HDE, &HAA, &H44, &HFF)
 FOLDERID_PublicMusic = iid
End Function

Public Function FOLDERID_PublicVideos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2400183A, CInt(&H6185), CInt(&H49FB), &HA2, &HD8, &H4A, &H39, &H2A, &H60, &H2B, &HA3)
 FOLDERID_PublicVideos = iid
End Function

Public Function FOLDERID_ResourceDir() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8AD10C31, CInt(&H2ADB), CInt(&H4296), &HA8, &HF7, &HE4, &H70, &H12, &H32, &HC9, &H72)
 FOLDERID_ResourceDir = iid
End Function

Public Function FOLDERID_LocalizedResourcesDir() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2A00375E, CInt(&H224C), CInt(&H49DE), &HB8, &HD1, &H44, &HD, &HF7, &HEF, &H3D, &HDC)
 FOLDERID_LocalizedResourcesDir = iid
End Function

Public Function FOLDERID_CommonOEMLinks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC1BAE2D0, CInt(&H10DF), CInt(&H4334), &HBE, &HDD, &H7A, &HA2, &HB, &H22, &H7A, &H9D)
 FOLDERID_CommonOEMLinks = iid
End Function

Public Function FOLDERID_CDBurning() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H9E52AB10, CInt(&HF80D), CInt(&H49DF), &HAC, &HB8, &H43, &H30, &HF5, &H68, &H78, &H55)
 FOLDERID_CDBurning = iid
End Function

Public Function FOLDERID_UserProfiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H762D272, CInt(&HC50A), CInt(&H4BB0), &HA3, &H82, &H69, &H7D, &HCD, &H72, &H9B, &H80)
 FOLDERID_UserProfiles = iid
End Function

Public Function FOLDERID_Playlists() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE92C1C7, CInt(&H837F), CInt(&H4F69), &HA3, &HBB, &H86, &HE6, &H31, &H20, &H4A, &H23)
 FOLDERID_Playlists = iid
End Function

Public Function FOLDERID_SamplePlaylists() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H15CA69B3, CInt(&H30EE), CInt(&H49C1), &HAC, &HE1, &H6B, &H5E, &HC3, &H72, &HAF, &HB5)
 FOLDERID_SamplePlaylists = iid
End Function

Public Function FOLDERID_SampleMusic() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB250C668, CInt(&HF57D), CInt(&H4EE1), &HA6, &H3C, &H29, &HE, &HE7, &HD1, &HAA, &H1F)
 FOLDERID_SampleMusic = iid
End Function

Public Function FOLDERID_SamplePictures() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HC4900540, CInt(&H2379), CInt(&H4C75), &H84, &H4B, &H64, &HE6, &HFA, &HF8, &H71, &H6B)
 FOLDERID_SamplePictures = iid
End Function

Public Function FOLDERID_SampleVideos() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H859EAD94, CInt(&H2E85), CInt(&H48AD), &HA7, &H1A, &H9, &H69, &HCB, &H56, &HA6, &HCD)
 FOLDERID_SampleVideos = iid
End Function

Public Function FOLDERID_PhotoAlbums() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H69D2CF90, CInt(&HFC33), CInt(&H4FB7), &H9A, &HC, &HEB, &HB0, &HF0, &HFC, &HB4, &H3C)
 FOLDERID_PhotoAlbums = iid
End Function

Public Function FOLDERID_Public() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDFDF76A2, CInt(&HC82A), CInt(&H4D63), &H90, &H6A, &H56, &H44, &HAC, &H45, &H73, &H85)
 FOLDERID_Public = iid
End Function

Public Function FOLDERID_ChangeRemovePrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDF7266AC, CInt(&H9274), CInt(&H4867), &H8D, &H55, &H3B, &HD6, &H61, &HDE, &H87, &H2D)
 FOLDERID_ChangeRemovePrograms = iid
End Function

Public Function FOLDERID_AppUpdates() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA305CE99, CInt(&HF527), CInt(&H492B), &H8B, &H1A, &H7E, &H76, &HFA, &H98, &HD6, &HE4)
 FOLDERID_AppUpdates = iid
End Function

Public Function FOLDERID_AddNewPrograms() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDE61D971, CInt(&H5EBC), CInt(&H4F02), &HA3, &HA9, &H6C, &H82, &H89, &H5E, &H5C, &H4)
 FOLDERID_AddNewPrograms = iid
End Function

Public Function FOLDERID_Downloads() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H374DE290, CInt(&H123F), CInt(&H4565), &H91, &H64, &H39, &HC4, &H92, &H5E, &H46, &H7B)
 FOLDERID_Downloads = iid
End Function

Public Function FOLDERID_PublicDownloads() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3D644C9B, CInt(&H1FB8), CInt(&H4F30), &H9B, &H45, &HF6, &H70, &H23, &H5F, &H79, &HC0)
 FOLDERID_PublicDownloads = iid
End Function

Public Function FOLDERID_SavedSearches() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7D1D3A04, CInt(&HDEBB), CInt(&H4115), &H95, &HCF, &H2F, &H29, &HDA, &H29, &H20, &HDA)
 FOLDERID_SavedSearches = iid
End Function

Public Function FOLDERID_QuickLaunch() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52A4F021, CInt(&H7B75), CInt(&H48A9), &H9F, &H6B, &H4B, &H87, &HA2, &H10, &HBC, &H8F)
 FOLDERID_QuickLaunch = iid
End Function

Public Function FOLDERID_Contacts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H56784854, CInt(&HC6CB), CInt(&H462B), &H81, &H69, &H88, &HE3, &H50, &HAC, &HB8, &H82)
 FOLDERID_Contacts = iid
End Function

Public Function FOLDERID_SidebarParts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HA75D362E, CInt(&H50FC), CInt(&H4FB7), &HAC, &H2C, &HA8, &HBE, &HAA, &H31, &H44, &H93)
 FOLDERID_SidebarParts = iid
End Function

Public Function FOLDERID_SidebarDefaultParts() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7B396E54, CInt(&H9EC5), CInt(&H4300), &HBE, &HA, &H24, &H82, &HEB, &HAE, &H1A, &H26)
 FOLDERID_SidebarDefaultParts = iid
End Function

Public Function FOLDERID_TreeProperties() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5B3749AD, CInt(&HB49F), CInt(&H49C1), &H83, &HEB, &H15, &H37, &HF, &HBD, &H48, &H82)
 FOLDERID_TreeProperties = iid
End Function

Public Function FOLDERID_PublicGameTasks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HDEBF2536, CInt(&HE1A8), CInt(&H4C59), &HB6, &HA2, &H41, &H45, &H86, &H47, &H6A, &HEA)
 FOLDERID_PublicGameTasks = iid
End Function

Public Function FOLDERID_GameTasks() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H54FAE61, CInt(&H4DD8), CInt(&H4787), &H80, &HB6, &H9, &H2, &H20, &HC4, &HB7, &H0)
 FOLDERID_GameTasks = iid
End Function

Public Function FOLDERID_SavedGames() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C5C32FF, CInt(&HBB9D), CInt(&H43B0), &HB5, &HB4, &H2D, &H72, &HE5, &H4E, &HAA, &HA4)
 FOLDERID_SavedGames = iid
End Function

Public Function FOLDERID_Games() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCAC52C1A, CInt(&HB53D), CInt(&H4EDC), &H92, &HD7, &H6B, &H2E, &H8A, &HC1, &H94, &H34)
 FOLDERID_Games = iid
End Function

Public Function FOLDERID_RecordedTV() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBD85E001, CInt(&H112E), CInt(&H431E), &H98, &H3B, &H7B, &H15, &HAC, &H9, &HFF, &HF1)
 FOLDERID_RecordedTV = iid
End Function

Public Function FOLDERID_SEARCH_MAPI() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H98EC0E18, CInt(&H2098), CInt(&H4D44), &H86, &H44, &H66, &H97, &H93, &H15, &HA2, &H81)
 FOLDERID_SEARCH_MAPI = iid
End Function

Public Function FOLDERID_SEARCH_CSC() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HEE32E446, CInt(&H31CA), CInt(&H4ABA), &H81, &H4F, &HA5, &HEB, &HD2, &HFD, &H6D, &H5E)
 FOLDERID_SEARCH_CSC = iid
End Function

Public Function FOLDERID_Links() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFB9D5E0, CInt(&HC6A9), CInt(&H404C), &HB2, &HB2, &HAE, &H6D, &HB6, &HAF, &H49, &H68)
 FOLDERID_Links = iid
End Function

Public Function FOLDERID_UsersFiles() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF3CE0F7C, CInt(&H4901), CInt(&H4ACC), &H86, &H48, &HD5, &HD4, &H4B, &H4, &HEF, &H8F)
 FOLDERID_UsersFiles = iid
End Function

Public Function FOLDERID_SearchHome() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H190337D1, CInt(&HB8CA), CInt(&H4121), &HA6, &H39, &H6D, &H47, &H2D, &H16, &H97, &H2A)
 FOLDERID_SearchHome = iid
End Function

Public Function FOLDERID_OriginalImages() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H2C36C0AA, CInt(&H5812), CInt(&H4B87), &HBF, &HD0, &H4C, &HD0, &HDF, &HB1, &H9B, &H39)
 FOLDERID_OriginalImages = iid
End Function

Public Function FOLDERID_HomeGroup() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H52528A6B, CInt(&HB9E3), CInt(&H4ADD), &HB6, &HD, &H58, &H8C, &H2D, &HBA, &H84, &H2D)
 FOLDERID_HomeGroup = iid
End Function
Public Function FOLDERID_AccountPictures() As UUID
'{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H8CA0B1, CInt(&H55B4), CInt(&H4C56), &HB8, &HA8, &H4D, &HE4, &HB2, &H99, &HD3, &HBE)
FOLDERID_AccountPictures = iid
End Function
Public Function FOLDERID_AppDataDesktop() As UUID
'{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB2C5E279, CInt(&H7ADD), CInt(&H439F), &HB2, &H8C, &HC4, &H1F, &HE1, &HBB, &HF6, &H72)
FOLDERID_AppDataDesktop = iid
End Function
Public Function FOLDERID_ApplicationShortcuts() As UUID
'{A3918781-E5F2-4890-B3D9-A7E54332328C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA3918781, CInt(&HE5F2), CInt(&H4890), &HB3, &HD9, &HA7, &HE5, &H43, &H32, &H32, &H8C)
FOLDERID_ApplicationShortcuts = iid
End Function
Public Function FOLDERID_AppsFolder() As UUID
'{1e87508d-89c2-42f0-8a7e-645a0f50ca58}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1E87508D, CInt(&H89C2), CInt(&H42F0), &H8A, &H7E, &H64, &H5A, &HF, &H50, &HCA, &H58)
FOLDERID_AppsFolder = iid
End Function
Public Function FOLDERID_CameraRoll() As UUID
'{AB5FB87B-7CE2-4F83-915D-550846C9537B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAB5FB87B, CInt(&H7CE2), CInt(&H4F83), &H91, &H5D, &H55, &H8, &H46, &HC9, &H53, &H7B)
FOLDERID_CameraRoll = iid
End Function
Public Function FOLDERID_DeviceMetadataStore() As UUID
'{5CE4A5E9-E4EB-479D-B89F-130C02886155}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CE4A5E9, CInt(&HE4EB), CInt(&H479D), &HB8, &H9F, &H13, &HC, &H2, &H88, &H61, &H55)
FOLDERID_DeviceMetadataStore = iid
End Function
Public Function FOLDERID_DocumentsLibrary() As UUID
'{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7B0DB17D, CInt(&H9CD2), CInt(&H4A93), &H97, &H33, &H46, &HCC, &H89, &H2, &H2E, &H7C)
FOLDERID_DocumentsLibrary = iid
End Function
Public Function FOLDERID_HomeGroupCurrentUser() As UUID
'{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9B74B6A3, CInt(&HDFD), CInt(&H4F11), &H9E, &H78, &H5F, &H78, &H0, &HF2, &HE7, &H72)
FOLDERID_HomeGroupCurrentUser = iid
End Function
Public Function FOLDERID_ImplicitAppShortcuts() As UUID
'{BCB5256F-79F6-4CEE-B725-DC34E402FD46}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCB5256F, CInt(&H79F6), CInt(&H4CEE), &HB7, &H25, &HDC, &H34, &HE4, &H2, &HFD, &H46)
FOLDERID_ImplicitAppShortcuts = iid
End Function
Public Function FOLDERID_Libraries() As UUID
'{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1B3EA5DC, CInt(&HB587), CInt(&H4786), &HB4, &HEF, &HBD, &H1D, &HC3, &H32, &HAE, &HAE)
FOLDERID_Libraries = iid
End Function
Public Function FOLDERID_MusicLibrary() As UUID
'{2112AB0A-C86A-4FFE-A368-0DE96E47012E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2112AB0A, CInt(&HC86A), CInt(&H4FFE), &HA3, &H68, &HD, &HE9, &H6E, &H47, &H1, &H2E)
FOLDERID_MusicLibrary = iid
End Function
Public Function FOLDERID_Objects3D() As UUID
'{31C0DD25-9439-4F12-BF41-7FF4EDA38722}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H31C0DD25, CInt(&H9439), CInt(&H4F12), &HBF, &H41, &H7F, &HF4, &HED, &HA3, &H87, &H22)
FOLDERID_Objects3D = iid
End Function
Public Function FOLDERID_PicturesLibrary() As UUID
'{A990AE9F-A03B-4E80-94BC-9912D7504104}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA990AE9F, CInt(&HA03B), CInt(&H4E80), &H94, &HBC, &H99, &H12, &HD7, &H50, &H41, &H4)
FOLDERID_PicturesLibrary = iid
End Function
Public Function FOLDERID_PublicLibraries() As UUID
'{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H48DAF80B, CInt(&HE6CF), CInt(&H4F4E), &HB8, &H0, &HE, &H69, &HD8, &H4E, &HE3, &H84)
FOLDERID_PublicLibraries = iid
End Function
Public Function FOLDERID_PublicRingtones() As UUID
'{E555AB60-153B-4D17-9F04-A5FE99FC15EC}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE555AB60, CInt(&H153B), CInt(&H4D17), &H9F, &H4, &HA5, &HFE, &H99, &HFC, &H15, &HEC)
FOLDERID_PublicRingtones = iid
End Function
Public Function FOLDERID_PublicUserTiles() As UUID
'{0482af6c-08f1-4c34-8c90-e17ec98b1e17}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H482AF6C, CInt(&H8F1), CInt(&H4C34), &H8C, &H90, &HE1, &H7E, &HC9, &H8B, &H1E, &H17)
FOLDERID_PublicUserTiles = iid
End Function
Public Function FOLDERID_RecordedTVLibrary() As UUID
'{1A6FDBA2-F42D-4358-A798-B74D745926C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1A6FDBA2, CInt(&HF42D), CInt(&H4358), &HA7, &H98, &HB7, &H4D, &H74, &H59, &H26, &HC5)
FOLDERID_RecordedTVLibrary = iid
End Function
Public Function FOLDERID_Ringtones() As UUID
'{C870044B-F49E-4126-A9C3-B52A1FF411E8}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC870044B, CInt(&HF49E), CInt(&H4126), &HA9, &HC3, &HB5, &H2A, &H1F, &HF4, &H11, &HE8)
FOLDERID_Ringtones = iid
End Function
Public Function FOLDERID_RoamedTileImages() As UUID
'{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAA8D5A5, CInt(&HF1D6), CInt(&H4259), &HBA, &HA8, &H78, &HE7, &HEF, &H60, &H83, &H5E)
FOLDERID_RoamedTileImages = iid
End Function
Public Function FOLDERID_RoamingTiles() As UUID
'{00BCFC5A-ED94-4e48-96A1-3F6217F21990}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCFC5A, CInt(&HED94), CInt(&H4E48), &H96, &HA1, &H3F, &H62, &H17, &HF2, &H19, &H90)
FOLDERID_RoamingTiles = iid
End Function
Public Function FOLDERID_SavedPictures() As UUID
'{3B193882-D3AD-4eab-965A-69829D1FB59F}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3B193882, CInt(&HD3AD), CInt(&H4EAB), &H96, &H5A, &H69, &H82, &H9D, &H1F, &HB5, &H9F)
FOLDERID_SavedPictures = iid
End Function
Public Function FOLDERID_SavedPicturesLibrary() As UUID
'{E25B5812-BE88-4bd9-94B0-29233477B6C3}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE25B5812, CInt(&HBE88), CInt(&H4BD9), &H94, &HB0, &H29, &H23, &H34, &H77, &HB6, &HC3)
FOLDERID_SavedPicturesLibrary = iid
End Function
Public Function FOLDERID_Screenshots() As UUID
'{b7bede81-df94-4682-a7d8-57a52620b86f}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HB7BEDE81, CInt(&HDF94), CInt(&H4682), &HA7, &HD8, &H57, &HA5, &H26, &H20, &HB8, &H6F)
FOLDERID_Screenshots = iid
End Function
Public Function FOLDERID_SearchHistory() As UUID
'{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HD4C3DB6, CInt(&H3A3), CInt(&H462F), &HA0, &HE6, &H8, &H92, &H4C, &H41, &HB5, &HD4)
FOLDERID_SearchHistory = iid
End Function
Public Function FOLDERID_SearchTemplates() As UUID
'{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7E636BFE, CInt(&HDFA9), CInt(&H4D5E), &HB4, &H56, &HD7, &HB3, &H98, &H51, &HD8, &HA9)
FOLDERID_SearchTemplates = iid
End Function
Public Function FOLDERID_SkyDrive() As UUID
'{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA52BBA46, CInt(&HE9E1), CInt(&H435F), &HB3, &HD9, &H28, &HDA, &HA6, &H48, &HC0, &HF6)
FOLDERID_SkyDrive = iid
End Function
Public Function FOLDERID_SkyDriveCameraRoll() As UUID
'{767E6811-49CB-4273-87C2-20F355E1085B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H767E6811, CInt(&H49CB), CInt(&H4273), &H87, &HC2, &H20, &HF3, &H55, &HE1, &H8, &H5B)
FOLDERID_SkyDriveCameraRoll = iid
End Function
Public Function FOLDERID_SkyDriveDocuments() As UUID
'{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H24D89E24, CInt(&H2F19), CInt(&H4534), &H9D, &HDE, &H6A, &H66, &H71, &HFB, &HB8, &HFE)
FOLDERID_SkyDriveDocuments = iid
End Function
Public Function FOLDERID_SkyDrivePictures() As UUID
'{339719B5-8C47-4894-94C2-D8F77ADD44A6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H339719B5, CInt(&H8C47), CInt(&H4894), &H94, &HC2, &HD8, &HF7, &H7A, &HDD, &H44, &HA6)
FOLDERID_SkyDrivePictures = iid
End Function
Public Function FOLDERID_UserPinned() As UUID
'{9E3995AB-1F9C-4F13-B827-48B24B6C7174}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9E3995AB, CInt(&H1F9C), CInt(&H4F13), &HB8, &H27, &H48, &HB2, &H4B, &H6C, &H71, &H74)
FOLDERID_UserPinned = iid
End Function
Public Function FOLDERID_UserProgramFiles() As UUID
'{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H5CD7AEE2, CInt(&H2219), CInt(&H4A67), &HB8, &H5D, &H6C, &H9C, &HE1, &H56, &H60, &HCB)
FOLDERID_UserProgramFiles = iid
End Function
Public Function FOLDERID_UserProgramFilesCommon() As UUID
'{BCBD3057-CA5C-4622-B42D-BC56DB0AE516}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBCBD3057, CInt(&HCA5C), CInt(&H4622), &HB4, &H2D, &HBC, &H56, &HDB, &HA, &HE5, &H16)
FOLDERID_UserProgramFilesCommon = iid
End Function
Public Function FOLDERID_UsersLibraries() As UUID
'{A302545D-DEFF-464b-ABE8-61C8648D939B}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA302545D, CInt(&HDEFF), CInt(&H464B), &HAB, &HE8, &H61, &HC8, &H64, &H8D, &H93, &H9B)
FOLDERID_UsersLibraries = iid
End Function
Public Function FOLDERID_VideosLibrary() As UUID
'{491E922F-5643-4AF4-A7EB-4E7A138D8174 }
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H491E922F, CInt(&H5643), CInt(&H4AF4), &HA7, &HEB, &H4E, &H7A, &H13, &H8D, &H81, &H74)
FOLDERID_VideosLibrary = iid
End Function
Public Function FOLDERID_RetailDemo() As UUID
'{12D4C69E-24AD-4923-BE19-31321C43A767}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H12D4C69E, CInt(&H24AD), CInt(&H4923), &HBE, &H19, &H31, &H32, &H1C, &H43, &HA7, &H67)
FOLDERID_RetailDemo = iid
End Function
Public Function FOLDERID_Device() As UUID
'{1C2AC1DC-4358-4B6C-9733-AF21156576F0}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H1C2AC1DC, CInt(&H4358), CInt(&H4B6C), &H97, &H33, &HAF, &H21, &H15, &H65, &H76, &HF0)
FOLDERID_Device = iid
End Function
Public Function FOLDERID_DevelopmentFiles() As UUID
'{DBE8E08E-3053-4BBC-B183-2A7B2B191E59}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDBE8E08E, CInt(&H3053), CInt(&H4BBC), &HB1, &H83, &H2A, &H7B, &H2B, &H19, &H1E, &H59)
FOLDERID_DevelopmentFiles = iid
End Function
Public Function FOLDERID_AppCaptures() As UUID
'{EDC0FE71-98D8-4F4A-B920-C8DC133CB165}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HEDC0FE71, CInt(&H98D8), CInt(&H4F4A), &HB9, &H20, &HC8, &HDC, &H13, &H3C, &HB1, &H65)
FOLDERID_AppCaptures = iid
End Function
Public Function FOLDERID_LocalDocuments() As UUID
'{f42ee2d3-909f-4907-8871-4c22fc0bf756}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HF42EE2D3, CInt(&H909F), CInt(&H4907), &H88, &H71, &H4C, &H22, &HFC, &HB, &HF7, &H56)
FOLDERID_LocalDocuments = iid
End Function
Public Function FOLDERID_LocalPictures() As UUID
'{0ddd015d-b06c-45d5-8c4c-f59713854639}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HDDD015D, CInt(&HB06C), CInt(&H45D5), &H8C, &H4C, &HF5, &H97, &H13, &H85, &H46, &H39)
FOLDERID_LocalPictures = iid
End Function
Public Function FOLDERID_LocalVideos() As UUID
'{35286a68-3c57-41a1-bbb1-0eae73d76c95}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H35286A68, CInt(&H3C57), CInt(&H41A1), &HBB, &HB1, &HE, &HAE, &H73, &HD7, &H6C, &H95)
FOLDERID_LocalVideos = iid
End Function
Public Function FOLDERID_LocalMusic() As UUID
'{a0c69a99-21c8-4671-8703-7934162fcf1d}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA0C69A99, CInt(&H21C8), CInt(&H4671), &H87, &H3, &H79, &H34, &H16, &H2F, &HCF, &H1D)
FOLDERID_LocalMusic = iid
End Function
Public Function FOLDERID_LocalDownloads() As UUID
'{7d83ee9b-2244-4e70-b1f5-5393042af1e4}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7D83EE9B, CInt(&H2244), CInt(&H4E70), &HB1, &HF5, &H53, &H93, &H4, &H2A, &HF1, &HE4)
FOLDERID_LocalDownloads = iid
End Function
Public Function FOLDERID_RecordedCalls() As UUID
'{2f8b40c2-83ed-48ee-b383-a1f157ec6f9a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H2F8B40C2, CInt(&H83ED), CInt(&H48EE), &HB3, &H83, &HA1, &HF1, &H57, &HEC, &H6F, &H9A)
FOLDERID_RecordedCalls = iid
End Function
Public Function FOLDERID_AllAppMods() As UUID
'{7ad67899-66af-43ba-9156-6aad42e6c596}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7AD67899, CInt(&H66AF), CInt(&H43BA), &H91, &H56, &H6A, &HAD, &H42, &HE6, &HC5, &H96)
FOLDERID_AllAppMods = iid
End Function
Public Function FOLDERID_CurrentAppMods() As UUID
'{3db40b20-2a30-4dbe-917e-771dd21dd099}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H3DB40B20, CInt(&H2A30), CInt(&H4DBE), &H91, &H7E, &H77, &H1D, &HD2, &H1D, &HD0, &H99)
FOLDERID_CurrentAppMods = iid
End Function
Public Function FOLDERID_AppDataDocuments() As UUID
'{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7BE16610, CInt(&H1F7F), CInt(&H44AC), &HBF, &HF0, &H83, &HE1, &H5F, &H2F, &HFC, &HA1)
FOLDERID_AppDataDocuments = iid
End Function
Public Function FOLDERID_AppDataFavorites() As UUID
'{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H7CFBEFBC, CInt(&HDE1F), CInt(&H45AA), &HB8, &H43, &HA5, &H42, &HAC, &H53, &H6C, &HC9)
FOLDERID_AppDataFavorites = iid
End Function
Public Function FOLDERID_AppDataProgramData() As UUID
'{559D40A3-A036-40FA-AF61-84CB430A4D34}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H559D40A3, CInt(&HA036), CInt(&H40FA), &HAF, &H61, &H84, &HCB, &H43, &HA, &H4D, &H34)
FOLDERID_AppDataProgramData = iid
End Function

Public Sub FreeKnownFolderDefinitionFields(pKFD As KNOWNFOLDER_DEFINITION)
Call CoTaskMemFree(pKFD.pszName)
Call CoTaskMemFree(pKFD.pszDescription)
Call CoTaskMemFree(pKFD.pszRelativePath)
Call CoTaskMemFree(pKFD.pszParsingName)
Call CoTaskMemFree(pKFD.pszToolTip)
Call CoTaskMemFree(pKFD.pszLocalizedName)
Call CoTaskMemFree(pKFD.pszIcon)
Call CoTaskMemFree(pKFD.pszSecurity)
End Sub
