VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Coupon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic

Option Explicit

Private mvarStartAirport As String
Private mvarEndAirport As String
Private mvarBIC As String
Private mvarDt As String
Private mvarStartTime As String
Private mvarAllowCnt As String
Private mvarAllowInd As String
Private mvarCouponStatus As String
Private mvarFlightNumber As String
Private mvarFareExpirary As String
Private mvarAirVendor As String
Private mvarFareBasis As String

Public Property Let AirVendor(ByVal vData As String)
    mvarAirVendor = vData
End Property

Public Property Get AirVendor() As String
    AirVendor = mvarAirVendor
End Property
Public Property Let FareBasis(ByVal vData As String)
    mvarFareBasis = vData
End Property

Public Property Get FareBasis() As String
    FareBasis = mvarFareBasis
End Property
Public Property Let FlightNumber(ByVal FlightNum As String)
    mvarFlightNumber = FlightNum
End Property

Public Property Get FlightNumber() As String
    FlightNumber = mvarFlightNumber
End Property
Public Property Let StartAirport(ByVal vData As String)
    mvarStartAirport = vData
End Property

Public Property Get StartAirport() As String
    StartAirport = mvarStartAirport
End Property

Public Property Let EndAirport(ByVal vData As String)
    mvarEndAirport = vData
End Property

Public Property Get EndAirport() As String
    EndAirport = mvarEndAirport
End Property

Public Property Let AllowCnt(ByVal vData As String)
    mvarAllowCnt = vData
End Property

Public Property Get AllowCnt() As String
    AllowCnt = mvarAllowCnt
End Property

Public Property Let AllowInd(ByVal vData As String)
    mvarAllowInd = vData
End Property

Public Property Get AllowInd() As String
    AllowInd = mvarAllowInd
End Property


Public Property Let BIC(ByVal vData As String)
    mvarBIC = vData
End Property

Public Property Get BIC() As String
    BIC = mvarBIC
End Property

Public Property Let Dt(ByVal vData As String)
    mvarDt = vData
End Property

Public Property Get Dt() As String
    Dt = mvarDt
End Property

Public Property Let StartTime(ByVal vData As String)
    mvarStartTime = vData
End Property

Public Property Get StartTime() As String
    StartTime = mvarStartTime
End Property


Public Property Let CouponStatus(ByVal vData As String)
    mvarCouponStatus = vData
End Property

Public Property Get CouponStatus() As String
    CouponStatus = mvarCouponStatus
End Property

Public Property Let FareExpire(ByVal vData As String)
    mvarFareExpirary = vData
End Property

Public Property Get FareExpire() As String
    FareExpire = mvarFareExpirary
End Property




