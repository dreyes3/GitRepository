VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EOTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCode As String
Private mvarAmount As Single

Public Property Let Amount(ByVal vData As Single)
    mvarAmount = vData
End Property

Public Property Get Amount() As Single
    Amount = mvarAmount
End Property

Public Property Let Code(ByVal vData As String)
    mvarCode = vData
End Property

Public Property Get Code() As String
    Code = mvarCode
End Property



