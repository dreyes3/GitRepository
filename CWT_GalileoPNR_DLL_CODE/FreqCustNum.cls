VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreqCustNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FreqCustNum"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mintPassengerNum As Integer 'local copy
Private mstrVendor As String 'local copy
Private mstrStatus As String 'local copy
Private mstrFreqCustNum As String 'local copy
Private mstrFieldType As String 'local copy
Private mstrCrossAccrual As String 'local copy

Public Property Let FieldType(ByVal TypeField As String)
    mstrFieldType = TypeField
End Property

Public Property Get FieldType() As String
    FieldType = mstrFieldType
End Property

Public Property Let FreqCustNum(ByVal FreqNumber As String)
    mstrFreqCustNum = FreqNumber
End Property

Public Property Get FreqCustNum() As String
    FreqCustNum = mstrFreqCustNum
End Property

Public Property Let Status(ByVal StatusCode As String)
    mstrStatus = StatusCode
End Property

Public Property Get Status() As String
    Status = mstrStatus
End Property

Public Property Let Vendor(ByVal VendorCode As String)
    mstrVendor = VendorCode
End Property

Public Property Get Vendor() As String
    Vendor = mstrVendor
End Property

Public Property Let PassengerNum(ByVal AbsPassengerNumber As Integer)
    mintPassengerNum = AbsPassengerNumber
End Property

Public Property Get PassengerNum() As Integer
    PassengerNum = mintPassengerNum
End Property

Public Property Let CrossAccrual(ByVal strField As String)
    mstrCrossAccrual = strField
End Property

Public Property Get CrossAccrual() As String
    CrossAccrual = mstrCrossAccrual
End Property

