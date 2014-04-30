VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Journey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)

Private mintFareIndex As Integer 'local copy
Private mintItem As Integer 'local copy
Private mlngJourneyTime As Long 'local copy

' Other variables
Private mcolAvailFlights As Collection
Public Sub AddAvailFlight(AirVCode As String, FlightNumber As String, _
    BIC As String, FromAptCode As String, ToAptCode As String, _
    DepartDtTm As Date, ArriveDtTm As Date, EquipCode As String, _
    FlyTime As Long, Optional CodeShareFlag As Boolean = False, _
    Optional NumberStops As Long = 0, Optional FareBasis As String = "")

Dim objAF As AvailFlight
Set objAF = New AvailFlight

With objAF
    .Airline = AirVCode
    .ArrDateTime = ArriveDtTm
    .CodeShare = CodeShareFlag
    .COS = BIC
    .DeptDateTime = DepartDtTm
    .Equip = EquipCode
    .FBC = FareBasis
    .FlightNum = FlightNumber
    .FlightTime = FlyTime
    .FromApt = FromAptCode
    .NumStops = NumberStops
    .ToApt = ToAptCode
End With

mcolAvailFlights.Add objAF
Set objAF = Nothing

    
End Sub


Public Property Let JourneyTime(ByVal JourneyTimeMinutes As Long)
    mlngJourneyTime = JourneyTimeMinutes
End Property

Public Property Get JourneyTime() As Long
    JourneyTime = mlngJourneyTime
End Property

Public Property Let Item(ByVal vData As Integer)
    mintItem = vData
End Property

Public Property Get Item() As Integer
Attribute Item.VB_UserMemId = 0
    Item = mintItem
End Property

Public Property Let FareIndex(ByVal vData As Integer)
    mintFareIndex = vData
End Property

Public Property Get FareIndex() As Integer
    FareIndex = mintFareIndex
End Property

Private Sub Class_Initialize()
    Set mcolAvailFlights = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set mcolAvailFlights = Nothing
End Sub

Public Property Get FlightCount() As Integer
    FlightCount = mcolAvailFlights.Count
End Property
