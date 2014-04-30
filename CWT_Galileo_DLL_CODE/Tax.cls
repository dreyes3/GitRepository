VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrTaxCode As String 'local copy
Private mcurAmount As Single 'local copy
Public Property Let Amount(ByVal TaxAmount As Single)
    mcurAmount = TaxAmount
End Property

Public Property Get Amount() As Single
    Amount = mcurAmount
End Property

Public Property Let TaxCode(ByVal Code As String)
    mstrTaxCode = Code
End Property

Public Property Get TaxCode() As String
    TaxCode = mstrTaxCode
End Property

