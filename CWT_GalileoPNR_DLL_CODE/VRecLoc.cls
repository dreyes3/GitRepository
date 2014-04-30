VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VRecLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarVendor As String 'local copy
Private mvarRecLoc As String 'local copy
Private mvarDateTime As Date 'local copy
Public Property Let DateTime(ByVal vData As Date)
    mvarDateTime = vData
End Property

Public Property Get DateTime() As Date
    DateTime = mvarDateTime
End Property

Public Property Let RecLoc(ByVal vData As String)
    mvarRecLoc = vData
End Property

Public Property Get RecLoc() As String
    RecLoc = mvarRecLoc
End Property

Public Property Let Vendor(ByVal vData As String)
    mvarVendor = vData
End Property

Public Property Get Vendor() As String
    Vendor = mvarVendor
End Property
