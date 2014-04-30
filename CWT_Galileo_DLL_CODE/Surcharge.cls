VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Surcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRelatedSegment As String 'local copy
Private mvarAmount As Single 'local copy
Private mvarCurrencyCode As String 'local copy
Public Property Let CurrencyCode(ByVal vData As String)
    mvarCurrencyCode = vData
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = mvarCurrencyCode
End Property

Public Property Let Amount(ByVal vData As Single)
    mvarAmount = vData
End Property

Public Property Get Amount() As Single
    Amount = mvarAmount
End Property

Public Property Let RelatedSegment(ByVal vData As String)
    mvarRelatedSegment = vData
End Property

Public Property Get RelatedSegment() As String
    RelatedSegment = mvarRelatedSegment
End Property

