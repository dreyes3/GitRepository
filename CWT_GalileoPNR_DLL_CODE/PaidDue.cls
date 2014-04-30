VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaidDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSegNum As Integer 'local copy
Private mvarProductType As String 'local copy
Private mvarSegDate As Date 'local copy
Private mvarAmount As Currency 'local copy
Private mvarFreeText As String 'local copy
Private mvarSegType As String 'local copy
'TUR
Private mvarCurrencyCode As String

Public Property Let SegType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SegType = 5
    mvarSegType = vData
End Property


Public Property Get SegType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SegType
    SegType = mvarSegType
End Property



Public Property Let FreeText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FreeText = 5
    mvarFreeText = vData
End Property


Public Property Get FreeText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FreeText
    FreeText = mvarFreeText
End Property

'TUR
Public Property Let CurrencyCode(ByVal vData As String)
    mvarCurrencyCode = vData
End Property

'TUR
Public Property Get CurrencyCode() As String
    CurrencyCode = mvarCurrencyCode
End Property

Public Property Let Amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    mvarAmount = vData
End Property


Public Property Get Amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    Amount = mvarAmount
End Property



Public Property Let SegDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SegDate = 5
    mvarSegDate = vData
End Property


Public Property Get SegDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SegDate
    SegDate = mvarSegDate
End Property



Public Property Let ProductType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProductType = 5
    mvarProductType = vData
End Property


Public Property Get ProductType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductType
    ProductType = mvarProductType
End Property



Public Property Let SegNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SegNum = 5
    mvarSegNum = vData
End Property


Public Property Get SegNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SegNum
    SegNum = mvarSegNum
End Property



