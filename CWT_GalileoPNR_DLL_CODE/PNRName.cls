VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PNRName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarGDSNum As String 'local copy
Private mvarPassengerNum As Integer 'local copy
Private mvarLastName As String 'local copy
Private mvarFirstName As String 'local copy
Private mvarRemark As String 'local copy
Private mvarPassengerType As String 'local copy
Public Property Let PassengerType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PassengerType = 5
    mvarPassengerType = vData
End Property


Public Property Get PassengerType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PassengerType
    PassengerType = mvarPassengerType
End Property



Public Property Let Remark(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Remark = 5
    mvarRemark = vData
End Property


Public Property Get Remark() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Remark
    Remark = mvarRemark
End Property



Public Property Let FirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarFirstName = vData
End Property


Public Property Get FirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    FirstName = mvarFirstName
End Property



Public Property Let LastName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastName = 5
    mvarLastName = vData
End Property


Public Property Get LastName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastName
    LastName = mvarLastName
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



Public Property Let GDSNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GDSNum = 5
    mvarGDSNum = vData
End Property


Public Property Get GDSNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GDSNum
    GDSNum = mvarGDSNum
End Property



