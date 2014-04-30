VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration

Option Explicit

Private mvarPlatingCarrier As String
Private mvarRouting As String
Private mvarBookingDate As Date
Private mvarBaseFare As Currency
Private mvarFareCurrency As String
Private mvarTax As Currency
Private mvarTaxCode As String
Private mvarConfirmationNum As String
Private mcolAirSeg As Collection

Public Property Let PlatingCarrier(ByVal vData As String)
    mvarPlatingCarrier = vData
End Property

Public Property Get PlatingCarrier() As String
    PlatingCarrier = mvarPlatingCarrier
End Property

Public Property Let Routing(ByVal vData As String)
    mvarRouting = vData
End Property

Public Property Get Routing() As String
    Routing = mvarRouting
End Property

Public Property Let BookingDate(ByVal vData As Date)
    mvarBookingDate = vData
End Property

Public Property Get BookingDate() As Date
    BookingDate = mvarBookingDate
End Property

Public Property Let BaseFare(ByVal vData As Currency)
    mvarBaseFare = vData
End Property

Public Property Get BaseFare() As Currency
    BaseFare = mvarBaseFare
End Property

Public Property Let FareCurrency(ByVal vData As String)
    mvarFareCurrency = vData
End Property

Public Property Get FareCurrency() As String
    FareCurrency = mvarFareCurrency
End Property

Public Property Let Tax(ByVal vData As Currency)
    mvarTax = vData
End Property

Public Property Get Tax() As Currency
    Tax = mvarTax
End Property
Public Property Let TaxCode(ByVal vData As String)
    mvarTaxCode = vData
End Property

Public Property Get TaxCode() As String
    TaxCode = mvarTaxCode
End Property

Public Property Let ConfirmationNum(ByVal vData As String)
    mvarConfirmationNum = vData
End Property

Public Property Get ConfirmationNum() As String
    ConfirmationNum = mvarConfirmationNum
End Property

Public Sub AddAirSeg(ByVal AirSeg As FareOptionSegment)
    mcolAirSeg.Add AirSeg
End Sub

Public Property Get AirSeg(Index As Integer) As FareOptionSegment
    Set AirSeg = mcolAirSeg.item(Index)
End Property

Public Property Get AirSegCount() As Integer
    AirSegCount = mcolAirSeg.Count
End Property

Private Sub Class_Initialize()
    Set mcolAirSeg = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolAirSeg = Nothing
End Sub

