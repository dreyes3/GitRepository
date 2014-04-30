VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAirport As String 'local copy
Private mvarAmount As Single 'local copy
Private mvarCurrencyCode As String 'local copy
Public Property Let CurrencyCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrencyCode = 5
    mvarCurrencyCode = vData
End Property


Public Property Get CurrencyCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrencyCode
    CurrencyCode = mvarCurrencyCode
End Property



Public Property Let Amount(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    mvarAmount = vData
End Property


Public Property Get Amount() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    Amount = mvarAmount
End Property



Public Property Let Airport(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Airport = 5
    mvarAirport = vData
End Property


Public Property Get Airport() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Airport
    Airport = mvarAirport
End Property



