VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AvailFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrAirline As String 'local copy
Private mstrFlightNum As String 'local copy
Private mstrCOS As String 'local copy
Private mstrToApt As String 'local copy
Private mstrFromApt As String 'local copy
Private mdtmDeptDateTime As Date 'local copy
Private mdtmArrDateTime As Date 'local copy
Private mstrEquip As String 'local copy
Private mintNumStops As Integer 'local copy
Private mblnCodeShare As Boolean 'local copy
Private mlngFlightTime As Long 'local copy
Private mstrFBC As String 'local copy

Public Property Let FBC(ByVal FBC As String)
    mstrFBC = FBC
End Property

Public Property Get FBC() As String
    FBC = mstrFBC
End Property

Public Property Let FlightTime(ByVal FlyTime As Long)
    mlngFlightTime = FlyTime
End Property

Public Property Get FlightTime() As Long
    FlightTime = mlngFlightTime
End Property

Public Property Let CodeShare(ByVal CodeShareFlag As Boolean)
    mblnCodeShare = CodeShareFlag
End Property

Public Property Get CodeShare() As Boolean
    CodeShare = mblnCodeShare
End Property

Public Property Let NumStops(ByVal NumOfStops As Integer)
    mintNumStops = NumOfStops
End Property

Public Property Get NumStops() As Integer
        NumStops = mintNumStops
End Property

Public Property Let Equip(ByVal EquipCode As String)
    mstrEquip = EquipCode
End Property

Public Property Get Equip() As String
    Equip = mstrEquip
End Property

Public Property Let ArrDateTime(ByVal ArrivDateTime As Date)
    mdtmArrDateTime = ArrivDateTime
End Property

Public Property Get ArrDateTime() As Date
    ArrDateTime = mdtmArrDateTime
End Property

Public Property Let DeptDateTime(ByVal DepartDateTime As Date)
    mdtmDeptDateTime = DepartDateTime
End Property

Public Property Get DeptDateTime() As Date
    DeptDateTime = mdtmDeptDateTime
End Property

Public Property Let FromApt(ByVal FromAptCode As String)
    mstrFromApt = FromAptCode
End Property

Public Property Get FromApt() As String
    FromApt = mstrFromApt
End Property

Public Property Let ToApt(ByVal ToAptCode As String)
    mstrToApt = ToAptCode
End Property

Public Property Get ToApt() As String
    ToApt = mstrToApt
End Property

Public Property Let COS(ByVal BIC As String)
    mstrCOS = BIC
End Property

Public Property Get COS() As String
    COS = mstrCOS
End Property

Public Property Let FlightNum(ByVal FlightNumber As String)
    mstrFlightNum = FlightNumber
End Property

Public Property Get FlightNum() As String
    FlightNum = mstrFlightNum
End Property

Public Property Let Airline(ByVal AirVendorCode As String)
    mstrAirline = AirVendorCode
End Property

Public Property Get Airline() As String
    Airline = mstrAirline
End Property
