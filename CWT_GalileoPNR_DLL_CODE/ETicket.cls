VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ETicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic

Option Explicit

Private mvarRoute As String
Private mvarRouteStart As String
Private mvarRouteEnd As String
Private mvarTiktNo As String
Private mvarFOP As String
Private mvarIssueDate As String
Private mvarFareCalc As String
Private mvarFares As String
Private mvarTaxes As String
Private mvarFees As String
Private mvarCharges As String
Private mvarTotal As String
Private mvarTourCode As String
Private mvarText As String
Private mvarPaxFName As String
Private mvarPaxLName As String
Private mvarPaxType As String
Private mcolCoupons As Collection

Public Property Let Route(ByVal vData As String)
    mvarRoute = vData
End Property

Public Property Get Route() As String
    Route = mvarRoute
End Property

Public Property Let RouteStart(ByVal vData As String)
    mvarRouteStart = vData
End Property

Public Property Get RouteStart() As String
    RouteStart = mvarRouteStart
End Property

Public Property Let RouteEnd(ByVal vData As String)
    mvarRouteEnd = vData
End Property

Public Property Get RouteEnd() As String
    RouteEnd = mvarRouteEnd
End Property

Public Property Let TiktNo(ByVal vData As String)
    mvarTiktNo = vData
End Property

Public Property Get TiktNo() As String
    TiktNo = mvarTiktNo
End Property

Public Property Let FOP(ByVal vData As String)
    mvarFOP = vData
End Property

Public Property Get FOP() As String
    FOP = mvarFOP
End Property

Public Property Let IssueDate(ByVal vData As String)
    mvarIssueDate = vData
End Property

Public Property Get IssueDate() As String
    IssueDate = mvarIssueDate
End Property

Public Property Let FareCalc(ByVal vData As String)
    mvarFareCalc = vData
End Property

Public Property Get FareCalc() As String
    FareCalc = mvarFareCalc
End Property

Public Property Let Fare(ByVal vData As String)
    mvarFares = vData
End Property

Public Property Get Fare() As String
    Fare = mvarFares
End Property

Public Property Let Taxes(ByVal vData As String)
    mvarTaxes = vData
End Property

Public Property Get Taxes() As String
    Taxes = mvarTaxes
End Property

Public Property Let Fees(ByVal vData As String)
    mvarFees = vData
End Property

Public Property Get Fees() As String
    Fees = mvarFees
End Property

Public Property Let Charges(ByVal vData As String)
    mvarCharges = vData
End Property

Public Property Get Charges() As String
    Charges = mvarCharges
End Property

Public Property Let Total(ByVal vData As String)
    mvarTotal = vData
End Property

Public Property Get Total() As String
    Total = mvarTotal
End Property

Public Property Let TourCode(ByVal vData As String)
    mvarTourCode = vData
End Property

Public Property Get TourCode() As String
    TourCode = mvarTourCode
End Property

Public Property Let EText(ByVal vData As String)
    mvarText = vData
End Property

Public Property Get EText() As String
    EText = mvarText
End Property

Public Property Let PaxFName(ByVal vData As String)
    mvarPaxFName = vData
End Property

Public Property Get PaxFName() As String
    PaxFName = mvarPaxFName
End Property

Public Property Let PaxLName(ByVal vData As String)
    mvarPaxLName = vData
End Property

Public Property Get PaxLName() As String
    PaxLName = mvarPaxLName
End Property


Public Property Let PaxType(ByVal vData As String)
    mvarPaxType = vData
End Property

Public Property Get PaxType() As String
    PaxType = mvarPaxType
End Property

Public Property Get CouponCount() As Long
    CouponCount = mcolCoupons.Count
End Property

Private Sub Class_Initialize()
    Set mcolCoupons = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolCoupons = Nothing
End Sub

Public Property Get Coupon(Index As Integer) As Coupon
    Set Coupon = mcolCoupons.Item(Index)
End Property
Public Sub AddCoupon(ByVal Coupon As Coupon)
    mcolCoupons.Add Coupon
End Sub

