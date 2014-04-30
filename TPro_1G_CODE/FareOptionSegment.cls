VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareOptionSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration

Private mvarDepCity As String
Private mvarArrCity As String
Private mvarDepTime As String
Private mvarArrTime As String
Private mvarFlightNum As String
Private mvarClass As String
Private mvarDepDate As String
Private mvarCarrier As String

Public Property Let DepDate(ByVal vData As String)
    mvarDepDate = vData
End Property

Public Property Get DepDate() As String
    DepDate = mvarDepDate
End Property

Public Property Let FlightNum(ByVal vData As String)
    mvarFlightNum = vData
End Property

Public Property Get FlightNum() As String

    FlightNum = mvarFlightNum
End Property

Public Property Let Class(ByVal vData As String)
    mvarClass = vData
End Property

Public Property Get Class() As String
    Class = mvarClass
End Property

Public Property Let DepCity(ByVal vData As String)
    mvarDepCity = vData
End Property

Public Property Get DepCity() As String
    DepCity = mvarDepCity
End Property

Public Property Let ArrCity(ByVal vData As String)
    mvarArrCity = vData
End Property

Public Property Get ArrCity() As String
    ArrCity = mvarArrCity
End Property
Public Property Let DepTime(ByVal vData As String)
    mvarDepTime = Format(vData, "0000")
End Property

Public Property Get DepTime() As String
    DepTime = mvarDepTime
End Property

Public Property Let ArrTime(ByVal vData As String)
    mvarArrTime = Format(vData, "0000")
End Property

Public Property Get ArrTime() As String
    ArrTime = mvarArrTime
End Property

Public Property Let Carrier(ByVal vData As String)
    mvarCarrier = vData
End Property

Public Property Get Carrier() As String
    Carrier = mvarCarrier
End Property

