VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AirSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Air Segment from Galileo Itinerary"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FareComponent"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mstrAirSegText As String
Private mintSegNum As Integer
Private mstrVendor As String
Private mstrFlightNum As String
Private mstrClass As String
Private mdtmDepDateTime As Date
Private mdtmArrDateTime As Date
Private mstrDepApt As String
Private mstrDepCityCode As String
Private mstrDepCountry As String
Private mstrDepCityName As String
Private mstrDepRegion As String
Private mstrDepIATAArea As String
Private mintDepDiffGMT As Integer
Private mstrArrApt As String
Private mstrArrCityCode As String
Private mstrArrCityName As String
Private mstrArrCountry As String
Private mstrArrRegion As String
Private mstrArrIATAArea As String
Private mintArrDiffGMT As Integer
Private mstrStatus As String 'local copy
Private mintNumberSeats As Integer 'local copy
Private mbolETicketEligible As Boolean 'local copy
Private mbolCodeShare As Boolean 'local copy
Private mstrOperatedBy As String 'local copy
Private mintFareComponent As Integer 'local copy
Private mbolSelectedForPricing As Boolean 'local copy
Private mbolConnection As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarFlown As Boolean 'local copy
Public Property Let Flown(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Flown = 5
    mvarFlown = vData
End Property

Public Property Get Flown() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Flown
    Flown = mvarFlown
End Property

Public Property Let Connection(ByVal Conx As Boolean)
    mbolConnection = Conx
End Property

Public Property Get Connection() As Boolean
    Connection = mbolConnection
End Property

Public Property Let SelectedForPricing(ByVal Selected As Boolean)
    mbolSelectedForPricing = Selected
End Property

Public Property Get SelectedForPricing() As Boolean
    SelectedForPricing = mbolSelectedForPricing
End Property

Public Property Let FareComponent(ByVal FareCompNumber As Integer)
    mintFareComponent = FareCompNumber
End Property

Public Property Get FareComponent() As Integer
    FareComponent = mintFareComponent
End Property

Public Property Let OperatedBy(ByVal OperatingCarrier As String)
    mstrOperatedBy = OperatingCarrier
End Property

Public Property Get OperatedBy() As String
    OperatedBy = mstrOperatedBy
End Property

Public Property Let CodeShareFlight(ByVal CodeShare As Boolean)
    mbolCodeShare = CodeShare
End Property

Public Property Get CodeShareFlight() As Boolean
    CodeShareFlight = mbolCodeShare
End Property

Public Property Let ETicketEligible(ByVal ETkt As Boolean)
    mbolETicketEligible = ETkt
End Property

Public Property Get ETicketEligible() As Boolean
    ETicketEligible = mbolETicketEligible
End Property

Public Property Let NumberSeats(ByVal SeatNum As Integer)
    mintNumberSeats = SeatNum
End Property

Public Property Get NumberSeats() As Integer
    NumberSeats = mintNumberSeats
End Property

Public Property Let Status(ByVal SegStatus As String)
    mstrStatus = SegStatus
End Property

Public Property Get Status() As String
    Status = mstrStatus
End Property

Public Property Get TextAirSeg() As String
    TextAirSeg = UCase(Format(CStr(mintSegNum), "@@. ") & mstrVendor & Format(mstrFlightNum, " @@@@") _
        & mstrClass & Format(mdtmDepDateTime, " ddmmmyy ") & mstrDepApt & mstrArrApt) & " " & mstrStatus & mintNumberSeats
End Property

Public Property Let SegNumber(ByVal SegNum As Integer)
    mintSegNum = SegNum
End Property

Public Property Get SegNumber() As Integer
    SegNumber = mintSegNum
End Property

Public Property Let Vendor(ByVal Carrier As String)
    mstrVendor = Carrier
End Property

Public Property Get Vendor() As String
    Vendor = mstrVendor
End Property

Public Property Let FlightNumber(ByVal FlightNum As String)
    mstrFlightNum = FlightNum
End Property

Public Property Get FlightNumber() As String
    FlightNumber = mstrFlightNum
End Property

Public Property Let Class(ByVal ClassOfSvc As String)
    mstrClass = ClassOfSvc
End Property

Public Property Get Class() As String
    Class = mstrClass
End Property

Public Property Let DepartDateTime(ByVal DepDate As Date)
    mdtmDepDateTime = DepDate
End Property

Public Property Get DepartDateTime() As Date
    DepartDateTime = mdtmDepDateTime
End Property

Public Property Let DepartAirport(ByVal DepApt As String)
    mstrDepApt = DepApt
End Property

Public Property Get DepartAirport() As String
    DepartAirport = mstrDepApt
End Property

Public Property Let DepartCityCode(ByVal DepCity As String)
    mstrDepCityCode = DepCity
End Property

Public Property Get DepartCityCode() As String
    DepartCityCode = mstrDepCityCode
End Property

Public Property Let DepartCityName(ByVal DepCity As String)
    mstrDepCityName = DepCity
End Property

Public Property Get DepartCityName() As String
    DepartCityName = mstrDepCityName
End Property

Public Property Let DepartCountry(ByVal DepCountry As String)
    mstrDepCountry = DepCountry
End Property

Public Property Get DepartCountry() As String
    DepartCountry = mstrDepCountry
End Property

Public Property Let DepartRegion(ByVal DepRegion As String)
    mstrDepRegion = DepRegion
End Property

Public Property Get DepartRegion() As String
    DepartRegion = mstrDepRegion
End Property

Public Property Let DepartIATAArea(ByVal DepIATAArea As String)
    mstrDepIATAArea = DepIATAArea
End Property

Public Property Get DepartIATAArea() As String
    DepartIATAArea = mstrDepIATAArea
End Property

Public Property Let DepartDiffGMT(ByVal DepDiffGMT As Integer)
    mintDepDiffGMT = DepDiffGMT
End Property

Public Property Get DepartDiffGMT() As Integer
    DepartDiffGMT = mintDepDiffGMT
End Property

Public Property Let ArriveDateTime(ByVal ArrDate As Date)
    mdtmArrDateTime = ArrDate
End Property

Public Property Get ArriveDateTime() As Date
    ArriveDateTime = mdtmArrDateTime
End Property

Public Property Let ArriveAirport(ByVal ArrApt As String)
    mstrArrApt = ArrApt
End Property

Public Property Get ArriveAirport() As String
    ArriveAirport = mstrArrApt
End Property

Public Property Let ArriveCityCode(ByVal ArrCity As String)
    mstrArrCityCode = ArrCity
End Property

Public Property Get ArriveCityCode() As String
    ArriveCityCode = mstrArrCityCode
End Property

Public Property Let ArriveCityName(ByVal ArrCity As String)
    mstrArrCityName = ArrCity
End Property

Public Property Get ArriveCityName() As String
    ArriveCityName = mstrArrCityName
End Property

Public Property Let ArriveCountry(ByVal ArrCountry As String)
    mstrArrCountry = ArrCountry
End Property

Public Property Get ArriveCountry() As String
    ArriveCountry = mstrArrCountry
End Property

Public Property Let ArriveRegion(ByVal ArrRegion As String)
    mstrArrRegion = ArrRegion
End Property

Public Property Get ArriveRegion() As String
    ArriveRegion = mstrArrRegion
End Property

Public Property Let ArriveIATAArea(ByVal ArrIATAArea As String)
    mstrArrIATAArea = ArrIATAArea
End Property

Public Property Get ArriveIATAArea() As String
    ArriveIATAArea = mstrArrIATAArea
End Property

Public Property Let ArriveDiffGMT(ByVal ArrDiffGMT As Integer)
    mintArrDiffGMT = ArrDiffGMT
End Property

Public Property Get ArriveDiffGMT() As Integer
    ArriveDiffGMT = mintArrDiffGMT
End Property

