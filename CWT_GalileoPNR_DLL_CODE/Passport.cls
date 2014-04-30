VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Passport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarPassportNum As String 'local copy
Private mvarDOB As Date 'local copy
Private mvarPassengerNum As Integer 'local copy
Private mvarExpDate As Date 'local copy
Private mvarGender As String 'local copy
Private mvarCitzCountryCode As String 'local copy
Private mvarResidenceCountryCode As String 'local copy
Public Property Let ResidenceCountryCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResidenceCountryCode = 5
    mvarResidenceCountryCode = vData
End Property


Public Property Get ResidenceCountryCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResidenceCountryCode
    ResidenceCountryCode = mvarResidenceCountryCode
End Property



Public Property Let CitzCountryCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CitzCountryCode = 5
    mvarCitzCountryCode = vData
End Property


Public Property Get CitzCountryCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CitzCountryCode
    CitzCountryCode = mvarCitzCountryCode
End Property



Public Property Let Gender(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gender = 5
    mvarGender = vData
End Property


Public Property Get Gender() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gender
    Gender = mvarGender
End Property



Public Property Let ExpDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExpDate = 5
    mvarExpDate = vData
End Property


Public Property Get ExpDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExpDate
    ExpDate = mvarExpDate
End Property



Public Property Let PassengerNum(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PassengerNum = 5
    mvarPassengerNum = vData
End Property


Public Property Get PassengerNum() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PassengerNum
    PassengerNum = mvarPassengerNum
End Property



Public Property Let DOB(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DOB = 5
    mvarDOB = vData
End Property


Public Property Get DOB() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DOB
    DOB = mvarDOB
End Property



Public Property Let PassportNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PassportNum = 5
    mvarPassportNum = vData
End Property


Public Property Get PassportNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PassportNum
    PassportNum = mvarPassportNum
End Property



