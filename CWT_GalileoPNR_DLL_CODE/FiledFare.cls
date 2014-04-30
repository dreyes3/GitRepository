VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FiledFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FiledFareTax"
Attribute VB_Ext_KEY = "Member1" ,"FiledFareSurcharge"
Attribute VB_Ext_KEY = "Member2" ,"FiledFareComponent"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTotalCurrency As String 'local copy
Private mvarTktTotalCurrency As String 'local copy
Private mvarTotAmount As Single 'local copy
Private mvarTktTotAmount As Single 'local copy
Private mvarJourneyNum As Integer 'local copy
Private mvarLastTktDate As Date 'local copy
Private mvarBaseCurrency As String 'local copy
Private mvarTktBaseCurrency As String 'local copy
Private mvarBaseAmount As Single 'local copy
Private mvarTktBaseAmount As Single 'local copy
Private mvarEquivCurrency As String 'local copy
Private mvarTktEquivCurrency As String 'local copy
Private mvarNetAmount As Single 'local copy
Private mvarSellAmount As Single 'local copy
Private mvarEquivAmount As Single 'local copy
Private mvarTktEquivAmount As Single 'local copy
Private mvarTaxTotal As Single
Private mvarTktTaxTotal As Single
Private mvarPFCTotal  As Single
Private mvarITNum As String
Private mvarPassengerNum As Long
Private mcolFFSegs As Collection
Private mcolFFComponents As Collection
Private mcolEndorsements As Collection
Private mcolTaxes As Collection
Private mcolTktTaxes As Collection
Private mcolPFC As Collection
Private mcolInfoMsg As Collection
Private mcolSurcharges As Collection
Private mvarJourneyType As String 'local copy
Private mvarFareConstructText As String 'local copy
Private mvarFareComponentCount As Integer 'local copy
Private mvarQuoteType As String 'local copy
Private mvarLastTicketDate As Date 'local copy
Private mvarISO As String 'local copy
Private mvarTCNum As String 'local copy
Private mvarPrivateFare As Boolean 'local copy
Private mvarSegSelectString As String 'local copy
Private mvarCommission As Single 'local copy
Private mvarFOP_CCExpireDate As Date 'local copy
Private mvarFOP_CCNum As String 'local copy
Private mvarFOPType As String 'local copy
Private mvarValidatingCarrier As String 'local copy
Private mvarTicketed As Boolean 'local copy
Private mvarCreatedDate As Date 'local copy
Private mvarCreatedBy As String 'local copy
'local variable(s) to hold property value(s)
Private mvarETktIndicator As Boolean 'local copy
Private mvarTicketNumber As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFOP_AddText As String 'local copy
Private mvarFOP_CCCode As String 'local copy
Private mvarASF As Single
Private mvarFareGuarCode As String
Private mvarTourCode As String
Private mvarCat35 As Boolean
Private mvarCommType As Boolean
Public Property Let TourCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FOP_CCCode = 5
    mvarTourCode = vData
End Property


Public Property Get TourCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FOP_CCCode
    TourCode = mvarTourCode
End Property
Public Property Let FareGuarCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FOP_CCCode = 5
    mvarFareGuarCode = vData
End Property


Public Property Get FareGuarCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FOP_CCCode
    FareGuarCode = mvarFareGuarCode
End Property

Public Property Let FOP_CCCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FOP_CCCode = 5
    mvarFOP_CCCode = vData
End Property


Public Property Get FOP_CCCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FOP_CCCode
    FOP_CCCode = mvarFOP_CCCode
End Property



Public Property Let FOP_AddText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FOP_AddText = 5
    mvarFOP_AddText = vData
End Property


Public Property Get FOP_AddText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FOP_AddText
    FOP_AddText = mvarFOP_AddText
End Property



Public Property Let TicketNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TicketNumber = 5
    mvarTicketNumber = vData
End Property


Public Property Get TicketNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TicketNumber
    TicketNumber = mvarTicketNumber
End Property



Public Property Let ETktIndicator(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ETktIndicator = 5
    mvarETktIndicator = vData
End Property


Public Property Get ETktIndicator() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ETktIndicator
    ETktIndicator = mvarETktIndicator
End Property



Public Property Let CreatedBy(ByVal vData As String)
    mvarCreatedBy = vData
End Property

Public Property Get CreatedBy() As String
    CreatedBy = mvarCreatedBy
End Property

Public Property Let CreatedDate(ByVal vData As Date)
    mvarCreatedDate = vData
End Property

Public Property Get CreatedDate() As Date
    CreatedDate = mvarCreatedDate
End Property

Public Property Let Ticketed(ByVal vData As Boolean)
    mvarTicketed = vData
End Property

Public Property Get Ticketed() As Boolean
    Ticketed = mvarTicketed
End Property

Public Property Let ValidatingCarrier(ByVal vData As String)
    mvarValidatingCarrier = vData
End Property

Public Property Get ValidatingCarrier() As String
    ValidatingCarrier = mvarValidatingCarrier
End Property

Public Property Let FOPType(ByVal vData As String)
    mvarFOPType = vData
End Property

Public Property Get FOPType() As String
    FOPType = mvarFOPType
End Property

Public Property Let FOP_CCNum(ByVal vData As String)
    mvarFOP_CCNum = vData
End Property

Public Property Get FOP_CCNum() As String
    FOP_CCNum = mvarFOP_CCNum
End Property

Public Property Let FOP_CCExpireDate(ByVal vData As Date)
    mvarFOP_CCExpireDate = vData
End Property

Public Property Get FOP_CCExpireDate() As Date
    FOP_CCExpireDate = mvarFOP_CCExpireDate
End Property

Public Property Let Commission(ByVal vData As Single)
    mvarCommission = vData
End Property

Public Property Get Commission() As Single
    Commission = mvarCommission
End Property
Public Property Let ASF(ByVal vData As Single)
    mvarASF = vData
End Property

Public Property Get ASF() As Single
    ASF = mvarASF
End Property
Public Property Let SegSelectString(ByVal vData As String)
    mvarSegSelectString = vData
End Property

Public Property Get SegSelectString() As String
    SegSelectString = mvarSegSelectString
End Property

Public Property Let PrivateFare(ByVal vData As Boolean)
    mvarPrivateFare = vData
End Property

Public Property Get PrivateFare() As Boolean
    PrivateFare = mvarPrivateFare
End Property

Public Property Let TCNum(ByVal vData As String)
    mvarTCNum = vData
End Property

Public Property Get TCNum() As String
    TCNum = mvarTCNum
End Property

Public Property Let ISO(ByVal vData As String)
    mvarISO = vData
End Property

Public Property Get ISO() As String
    ISO = mvarISO
End Property

Public Property Let LastTicketDate(ByVal vData As Date)
    mvarLastTicketDate = vData
End Property

Public Property Get LastTicketDate() As Date
    LastTicketDate = mvarLastTicketDate
End Property

Public Property Let QuoteType(ByVal vData As String)
    mvarQuoteType = vData
End Property

Public Property Get QuoteType() As String
    QuoteType = mvarQuoteType
End Property

Public Function FiledFareComponent(ByVal Index As Long) As FiledFareComponent
    Set FiledFareComponent = mcolFFComponents.Item(Index)
End Function

Public Sub FareComponentAdd(ByVal NewFareComponent As FiledFareComponent)
    mcolFFComponents.Add NewFareComponent
    Set NewFareComponent = Nothing
End Sub

Public Property Get FareComponentCount() As Long
    FareComponentCount = mcolFFComponents.Count
End Property

Public Property Let FareConstructText(ByVal FareConstructText As String)
    mvarFareConstructText = FareConstructText
End Property

Public Property Get FareConstructText() As String
    FareConstructText = mvarFareConstructText
End Property

Public Property Let JourneyType(ByVal vData As String)
    mvarJourneyType = vData
End Property

Public Property Get JourneyType() As String
    JourneyType = mvarJourneyType
End Property

Public Property Let EquivAmount(ByVal vData As Single)
    mvarEquivAmount = vData
End Property

Public Property Get EquivAmount() As Single
    EquivAmount = mvarEquivAmount
End Property

Public Property Let EquivCurrency(ByVal vData As String)
    mvarEquivCurrency = vData
End Property

Public Property Get EquivCurrency() As String
    EquivCurrency = mvarEquivCurrency
End Property
Public Property Let TktEquivAmount(ByVal vData As Single)
    mvarTktEquivAmount = vData
End Property

Public Property Get TktEquivAmount() As Single
    TktEquivAmount = mvarTktEquivAmount
End Property

Public Property Let TktEquivCurrency(ByVal vData As String)
    mvarTktEquivCurrency = vData
End Property

Public Property Get TktEquivCurrency() As String
    TktEquivCurrency = mvarTktEquivCurrency
End Property


Public Property Get TaxTotal() As Single
    TaxTotal = mvarTaxTotal
End Property
Public Property Get TktTaxTotal() As Single
    TktTaxTotal = mvarTktTaxTotal
End Property

Public Property Let BaseAmount(ByVal vData As Single)
    mvarBaseAmount = vData
End Property
Public Property Let TktBaseAmount(ByVal vData As Single)
    mvarTktBaseAmount = vData
End Property

Public Property Get BaseAmount() As Single
    BaseAmount = mvarBaseAmount
End Property
Public Property Get TktBaseAmount() As Single
    TktBaseAmount = mvarTktBaseAmount
End Property

Public Property Let NetAmount(ByVal vData As Single)
    mvarNetAmount = vData
End Property

Public Property Get NetAmount() As Single
    NetAmount = mvarNetAmount
End Property

Public Property Let SellAmount(ByVal vData As Single)
    mvarSellAmount = vData
End Property

Public Property Get SellAmount() As Single
    SellAmount = mvarSellAmount
End Property

Public Property Let PassengerNum(ByVal vData As Long)
    mvarPassengerNum = vData
End Property

Public Property Get PassengerNum() As Long
    PassengerNum = mvarPassengerNum
End Property

Public Property Let BaseCurrency(ByVal vData As String)
    mvarBaseCurrency = vData
End Property

Public Property Get BaseCurrency() As String
    BaseCurrency = mvarBaseCurrency
End Property
Public Property Let TktBaseCurrency(ByVal vData As String)
    mvarTktBaseCurrency = vData
End Property

Public Property Get TktBaseCurrency() As String
    TktBaseCurrency = mvarTktBaseCurrency
End Property


Public Property Let LastTktDate(ByVal vData As Date)
    mvarLastTktDate = vData
End Property

Public Property Get LastTktDate() As Date
    LastTktDate = mvarLastTktDate
End Property

Public Function SegmentData(ByVal Index As Variant) As FiledFareSegment
    Set SegmentData = mcolFFSegs.Item(Index)
End Function

Public Function Tax(ByVal Index As Variant) As FiledFareTax
    Set Tax = mcolTaxes.Item(Index)
End Function
Public Function TktTax(ByVal Index As Variant) As FiledFareTax
    Set TktTax = mcolTktTaxes.Item(Index)
End Function

Public Sub TaxAdd(ByVal NewTax As FiledFareTax)
    mcolTaxes.Add NewTax
    mvarTaxTotal = mvarTaxTotal + NewTax.Amount
End Sub
Public Sub TktTaxAdd(ByVal NewTax As FiledFareTax)
    mcolTktTaxes.Add NewTax
    mvarTktTaxTotal = mvarTktTaxTotal + NewTax.Amount
End Sub

Public Property Get TaxCount() As Long
    TaxCount = mcolTaxes.Count
End Property
Public Property Get TktTaxCount() As Long
    TktTaxCount = mcolTktTaxes.Count
End Property

Public Sub PFCAdd(ByVal NewPFC As PFC)
mcolPFC.Add NewPFC
mvarPFCTotal = mvarPFCTotal + NewPFC.Amount
End Sub

Public Property Get PFCCount() As Long
    PFCCount = mcolPFC.Count
End Property

Public Property Get PFCTotal() As Single
    PFCTotal = mvarPFCTotal
End Property
Public Sub FareSegAdd(ByVal NewSegment As FiledFareSegment)
    mcolFFSegs.Add NewSegment
End Sub

Public Function FareSeg(ByVal Index As Long) As FiledFareSegment
    Set FareSeg = mcolFFSegs.Item(Index)
End Function

Public Sub SurchargeAdd(ByVal NewSurcharge As FiledFareSurcharge)
    mcolSurcharges.Add NewSurcharge
    Set NewSurcharge = Nothing
End Sub

Public Function FiledFareSurcharge(ByVal Index As Long) As FiledFareSurcharge
    Set FiledFareSurcharge = mcolSurcharges.Item(Index)
End Function

Public Property Get SurchargeCount() As Long
    SurchargeCount = mcolSurcharges.Count
End Property

Public Property Get FareSegCount() As Long
    FareSegCount = mcolFFSegs.Count
End Property

Public Sub InfoMsgAdd(ByVal Message As String)
    mcolInfoMsg.Add Message
    
End Sub

Public Property Get InfoMsgCount() As Long
    InfoMsgCount = mcolInfoMsg.Count
End Property
Public Property Let JourneyNum(ByVal vData As Integer)
    mvarJourneyNum = vData
End Property

Public Property Get JourneyNum() As Integer
    JourneyNum = mvarJourneyNum
End Property

Public Property Let ITNum(ByVal vData As String)
    mvarITNum = vData
End Property

Public Property Get ITNum() As String
    ITNum = mvarITNum
End Property

Public Property Let TotAmount(ByVal vData As Single)
    mvarTotAmount = vData
End Property

Public Property Get TotAmount() As Single
    TotAmount = mvarTotAmount
End Property

Public Property Let TotalCurrency(ByVal vData As String)
    mvarTotalCurrency = vData
End Property

Public Property Get TotalCurrency() As String
    TotalCurrency = mvarTotalCurrency
End Property
Public Property Let TktTotAmount(ByVal vData As Single)
    mvarTktTotAmount = vData
End Property

Public Property Get TktTotAmount() As Single
    TktTotAmount = mvarTktTotAmount
End Property

Public Property Let TktTotalCurrency(ByVal vData As String)
    mvarTktTotalCurrency = vData
End Property

Public Property Get TktTotalCurrency() As String
    TktTotalCurrency = mvarTktTotalCurrency
End Property


Private Sub Class_Initialize()
    Set mcolTaxes = New Collection
    Set mcolTktTaxes = New Collection
    Set mcolSurcharges = New Collection
    Set mcolFFSegs = New Collection
    Set mcolInfoMsg = New Collection
    Set mcolFFComponents = New Collection
    Set mcolPFC = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolTaxes = Nothing
    Set mcolTktTaxes = Nothing
    Set mcolSurcharges = Nothing
    Set mcolFFSegs = Nothing
    Set mcolInfoMsg = Nothing
    Set mcolFFComponents = Nothing
    Set mcolPFC = Nothing
End Sub

Public Property Let Cat35(ByVal vData As Boolean)
    mvarCat35 = vData
End Property

Public Property Get Cat35() As Boolean
    Cat35 = mvarCat35
End Property

Public Property Let CommType(ByVal vData As Boolean)
    mvarCommType = vData
End Property

Public Property Get CommType() As Boolean
    CommType = mvarCommType
End Property

