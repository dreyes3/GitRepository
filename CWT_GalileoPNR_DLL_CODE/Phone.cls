VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCityCode As String 'local copy
Private mvarPhoneType As String 'local copy
Private mvarPhoneNum As String 'local copy
Private mvarItemNum As Integer 'local copy
Public Property Let ItemNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemNum = 5
    mvarItemNum = vData
End Property


Public Property Get ItemNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemNum
    ItemNum = mvarItemNum
End Property



Public Property Let PhoneNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhoneNum = 5
    mvarPhoneNum = vData
End Property


Public Property Get PhoneNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhoneNum
    PhoneNum = mvarPhoneNum
End Property



Public Property Let PhoneType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PhoneType = 5
    mvarPhoneType = vData
End Property


Public Property Get PhoneType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PhoneType
    PhoneType = mvarPhoneType
End Property



Public Property Let CityCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CityCode = 5
    mvarCityCode = vData
End Property


Public Property Get CityCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CityCode
    CityCode = mvarCityCode
End Property



