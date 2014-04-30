VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrFBC As String 'local copy
Private mdtmNVB As Date 'local copy
Private mdtmNVA As Date 'local copy
Private mdtmDepDate As Date
Private mstrTD As String 'local copy
Private mstrBagInfo As String 'local copy
Private mbolStopover As Boolean 'local copy
Private mstrInfoText As String 'local copy
'local variable(s) to hold property value(s)
Private mvarVendor As String 'local copy
Private mvarFlightNum As String 'local copy
Private mvarDepCityCode As String 'local copy
Private mvarArrCityCode As String 'local copy
Private mvarCOS As String 'local copy
Private mvarSegmentType As String
Private mstrOverridePFBC As String
Public Property Let COS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.COS = 5
    mvarCOS = vData
End Property


Public Property Get COS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.COS
    COS = mvarCOS
End Property



Public Property Let ArrCityCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ArrCityCode = 5
    mvarArrCityCode = vData
End Property


Public Property Get ArrCityCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ArrCityCode
    ArrCityCode = mvarArrCityCode
End Property



Public Property Let DepCityCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DepCityCode = 5
    mvarDepCityCode = vData
End Property


Public Property Get DepCityCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DepCityCode
    DepCityCode = mvarDepCityCode
End Property



Public Property Let SegType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlightNum = 5
    mvarSegmentType = vData
End Property


Public Property Get SegType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlightNum
    SegType = mvarSegmentType
End Property

Public Property Let FlightNum(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FlightNum = 5
    mvarFlightNum = vData
End Property


Public Property Get FlightNum() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FlightNum
    FlightNum = mvarFlightNum
End Property

Public Property Let Vendor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vendor = 5
    mvarVendor = vData
End Property


Public Property Get Vendor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vendor
    Vendor = mvarVendor
End Property



Public Property Let InfoText(ByVal TextData As String)
    mstrInfoText = TextData
End Property

Public Property Get InfoText() As String
    InfoText = mstrInfoText
End Property

Public Property Let Stopover(ByVal StopInd As Boolean)
    mbolStopover = StopInd
End Property

Public Property Get Stopover() As Boolean
    Stopover = mbolStopover
End Property

Public Property Let BagInfo(ByVal BagData As String)
    mstrBagInfo = BagData
End Property

Public Property Get BagInfo() As String
    BagInfo = mstrBagInfo
End Property

Public Property Let TD(ByVal TicketDesig As String)
    mstrTD = TicketDesig
End Property

Public Property Get TD() As String
    TD = mstrTD
End Property

Public Property Let NVA(ByVal NotValidAfter As Date)
    mdtmNVA = NotValidAfter
End Property

Public Property Get NVA() As Date
    NVA = mdtmNVA
End Property

Public Property Let NVB(ByVal NotValidBefore As Date)
    mdtmNVB = NotValidBefore
End Property

Public Property Get NVB() As Date
    NVB = mdtmNVB
End Property

Public Property Let FBC(ByVal FareBasisCode As String)
    mstrFBC = FareBasisCode
End Property

Public Property Get FBC() As String
    FBC = mstrFBC
End Property
Public Property Let OverridePFBC(ByVal OverridePFBC As String)
    mstrOverridePFBC = OverridePFBC
End Property

Public Property Get OverridePFBC() As String
    OverridePFBC = mstrOverridePFBC
End Property

Public Property Let DepDate(ByVal DepartureDate As Date)
    mdtmDepDate = DepartureDate
End Property

Public Property Get DepDate() As Date
    DepDate = mdtmDepDate
End Property
