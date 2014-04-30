VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VendorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'local variable(s) to hold property value(s)
Private mcolVendors As Collection






'Public Function VendorInfo(ByVal Index As Long) As Vendor
'    Set VendorInfo = mcolVendors.Item(Index)
'End Function

Public Sub VendorAdd(strEmail As String, strFax As String)
Dim objNewVendor As Vendor

Set objNewVendor = New Vendor
With objNewVendor
    .EmailAddr = strEmail
    .FaxNo = strFax
End With

mcolVendors.Add objNewVendor


'    mcolVendors.Add NewVendor
'    Set NewVendor = Nothing
End Sub

Public Property Get VendorCount() As Long
    VendorCount = mcolVendors.Count
End Property



Private Sub Class_Initialize()
    Set mcolVendors = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolVendors = Nothing
End Sub

