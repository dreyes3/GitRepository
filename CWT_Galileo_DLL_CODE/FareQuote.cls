VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Surcharge"
Attribute VB_Ext_KEY = "Member1" ,"Tax"
Attribute VB_Ext_KEY = "Member2" ,"FareComponent"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTotalCurrency As String 'local copy
Private mvarTktTotalCurrency As String
Private mvarTotAmount As Single 'local copy
Private mvarTktTotAmount As Single
Private mvarJourneyNum As Integer 'local copy
Private mvarLastTktDate As Date 'local copy
Private mvarBaseCurrency As String 'local copy
Private mvarTktBaseCurrency As String
Private mvarBaseAmount As Single 'local copy
Private mvarTktBaseAmount As Single
Private mvarCommission As Single 'local copy
Private mvarCommissionPt As Single
Private mvarEquivCurrency As String 'local copy
Private mvarTktEquivCurrency As String
Private mvarNetAmount As Single 'local copy
Private mvarSellAmount As Single 'local copy
Private mvarLowFare As Single
Private mvarMerchAmt As Single
Private mvarDiscountAmt As Single
Private mvarEquivAmount As Single 'local copy
Private mvarTktEquivAmount As Single
Private mvarTaxTotal As Single
Private mvarTktTaxTotal As Single
Private mvarITNum As String
Private mcolFareSegs As Collection
Private mcolFareComponents As Collection
Private mcolTaxes As Collection
Private mcolTktTaxes As Collection
Private mcolInfoMsg As Collection
Private mcolSurcharges As Collection
Private mvarJourneyType As String 'local copy
Private mvarFareConstructText As String 'local copy
Private mvarFareComponentCount As Integer 'local copy
Private mvarQuoteType As String 'local copy
Private mvarUnableToQuote As Boolean
Private mvarLastTicketDate As Date 'local copy
Private mvarISO As String 'local copy
Private mvarTCNum As String 'local copy
Private mvarPrivateFare As Boolean 'local copy
Private mvarROE As Single 'local copy
Private mvarSegmentString As String
Private mvarPIC As String
Private mvarPxNum As Integer
Private mbolSegmentSelected As Boolean
Private mvarHighFareComponent As Long 'local copy
Private mvarHPFApplies As Boolean 'local copy
Private mvarStoreFare As Boolean
Private mvarTransactionFee As Single 'local copy
Private mvarHiddenComm As Single 'local copy
Private mvarPFAccountCode As String 'local copy
Private mvarPFFareType As String
Private mvarNRCC As Boolean
Private mvarNoFare As Boolean
Private mbolLog As Boolean
Private gobjLog As CWT_AppLog.AppLog
Private mvarClientType As String
Private mvarComPct As Single
Private mvarDiscountPct As Single
Private mvarMerchPct As Single
Private mvarOverrideTF As Boolean
Private mvarInternational As Boolean
Private mvarOverrideConx As Boolean
Private mvarFQPCC As String
Private mvarPlatCarrier As String
Private mvarCat35 As Boolean
Private mvarCommType As Boolean
'230108
Private mvarFuelSurcharge As Single
Private mvarPTkt As Boolean
Private mvarRuleID As String
'Added by JiYong - Add in actual net amount
Private mvarActualNetAmount As Single

Public Property Let FuelSurcharge(ByVal vData As Single)
    mvarFuelSurcharge = vData
End Property

Public Property Get FuelSurcharge() As Single
    FuelSurcharge = mvarFuelSurcharge
End Property

Public Property Let FQPCC(ByVal vData As String)
   mvarFQPCC = vData
End Property

Public Property Get FQPCC() As String
    FQPCC = mvarFQPCC
End Property
Public Property Let OverrideConx(ByVal vData As Boolean)
    mvarOverrideConx = vData
End Property

Public Property Get OverrideConx() As Boolean
    OverrideConx = mvarOverrideConx
End Property
Public Property Let ClientType(ByVal vData As String)
   mvarClientType = vData
End Property

Public Property Get ClientType() As String
    ClientType = mvarClientType
End Property
Public Property Let DiscountPct(ByVal vData As Single)
    mvarDiscountPct = vData
End Property

Public Property Get DiscountPct() As Single
    DiscountPct = mvarDiscountPct
End Property
Public Property Let ComPct(ByVal vData As Single)
    mvarComPct = vData
End Property

Public Property Get ComPct() As Single
    ComPct = mvarComPct
End Property
Public Property Let MerchPct(ByVal vData As Single)
    mvarMerchPct = vData
End Property

Public Property Get MerchPct() As Single
    MerchPct = mvarMerchPct
End Property
Public Property Let International(ByVal vData As Boolean)
    mvarInternational = vData
End Property

Public Property Get International() As Boolean
    International = mvarInternational
End Property

Public Property Let OverrideTF(ByVal vData As Boolean)
    mvarOverrideTF = vData
End Property

Public Property Get OverrideTF() As Boolean
    OverrideTF = mvarOverrideTF
End Property
Public Property Let NoFare(ByVal vData As Boolean)
    mvarNoFare = vData
End Property

Public Property Get NoFare() As Boolean
    NoFare = mvarNoFare
End Property
Public Property Let NRCC(ByVal vData As Boolean)
    mvarNRCC = vData
End Property

Public Property Get NRCC() As Boolean
    NRCC = mvarNRCC
End Property
Public Property Let PFFareType(ByVal vData As String)
    mvarPFFareType = vData
End Property

Public Property Get PFFareType() As String
    PFFareType = mvarPFFareType
End Property
Public Property Let PFAccountCode(ByVal vData As String)
    mvarPFAccountCode = vData
End Property

Public Property Get PFAccountCode() As String
    PFAccountCode = mvarPFAccountCode
End Property

Public Property Let LowFare(ByVal vData As Single)
    mvarLowFare = vData
End Property

Public Property Get LowFare() As Single
    LowFare = mvarLowFare
End Property
Public Property Let MerchAmt(ByVal vData As Single)
    mvarMerchAmt = vData
End Property

Public Property Get MerchAmt() As Single
    MerchAmt = mvarMerchAmt
End Property

Public Property Let DiscountAmt(ByVal vData As Single)
    mvarDiscountAmt = vData
End Property

Public Property Get DiscountAmt() As Single
    DiscountAmt = mvarDiscountAmt
End Property

Public Property Let TransactionFee(ByVal vData As Single)
    mvarTransactionFee = vData
End Property

Public Property Get TransactionFee() As Single
    TransactionFee = mvarTransactionFee
End Property
Public Property Let HiddenComm(ByVal vData As Single)
    mvarHiddenComm = vData
End Property

Public Property Get HiddenComm() As Single
    HiddenComm = mvarHiddenComm
End Property

Public Property Let HPFApplies(ByVal vData As Boolean)
    mvarHPFApplies = vData
End Property

Public Property Get HPFApplies() As Boolean
    HPFApplies = mvarHPFApplies
End Property

Public Property Let HighFareComponent(ByVal vData As Long)
    mvarHighFareComponent = vData
End Property

Public Property Get HighFareComponent() As Long
    HighFareComponent = mvarHighFareComponent
End Property

Public Property Get ROE() As Single
    ROE = mvarROE
End Property

Public Property Let PrivateFare(ByVal vData As Boolean)
    mvarPrivateFare = vData
End Property

Public Property Get PrivateFare() As Boolean
    PrivateFare = mvarPrivateFare
End Property

Public Property Get StoreFare() As Boolean
    StoreFare = mvarStoreFare
End Property

Public Property Let StoreFare(ByVal vData As Boolean)
    mvarStoreFare = vData
End Property

Public Property Let SegmentSelected(ByVal vData As Boolean)
    mbolSegmentSelected = vData
End Property

Public Property Get SegmentSelected() As Boolean
   SegmentSelected = mbolSegmentSelected
End Property

Public Property Let TCNum(ByVal vData As String)
    mvarTCNum = vData
End Property

Public Property Get TCNum() As String
    TCNum = mvarTCNum
End Property

Public Property Let PIC(ByVal vData As String)
    mvarPIC = vData
End Property

Public Property Get PIC() As String
    PIC = mvarPIC
End Property

Public Property Let PxNum(ByVal vData As Integer)
    mvarPxNum = vData
End Property

Public Property Get PxNum() As Integer
    PxNum = mvarPxNum
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

Public Property Let UnableToQuote(ByVal vData As Boolean)
    mvarUnableToQuote = vData
End Property

Public Property Get UnableToQuote() As Boolean
    UnableToQuote = mvarUnableToQuote
End Property

Public Function FareComponent(ByVal Index As Long) As FareComponent
    Set FareComponent = mcolFareComponents.Item(Index)
End Function

Public Sub FareComponentAdd(ByVal NewFareComponent As FareComponent)
    mcolFareComponents.Add NewFareComponent
    Set NewFareComponent = Nothing
End Sub

Public Property Get FareComponentCount() As Long
    FareComponentCount = mcolFareComponents.Count
End Property

Public Property Let FareConstructText(ByVal FareConstructText As String)
Dim lngBegStr As Long
    
    mvarFareConstructText = FareConstructText
    
    lngBegStr = InStr(FareConstructText, "END ROE") + 7
    If lngBegStr > 7 Then
        mvarROE = CSng(Mid(FareConstructText, lngBegStr))
    Else
        mvarROE = 1
    End If
    
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
Public Property Let TktEquivAmount(ByVal vData As Single)
    mvarTktEquivAmount = vData
End Property

Public Property Get TktEquivAmount() As Single
    TktEquivAmount = mvarTktEquivAmount
End Property

Public Property Let EquivCurrency(ByVal vData As String)
    mvarEquivCurrency = vData
End Property

Public Property Get EquivCurrency() As String
    EquivCurrency = mvarEquivCurrency
End Property
Public Property Let TktEquivCurrency(ByVal vData As String)
    mvarTktEquivCurrency = vData
End Property

Public Property Get TktEquivCurrency() As String
    TktEquivCurrency = mvarTktEquivCurrency
End Property

Public Property Let TaxTotal(ByVal vData As Single)
    mvarTaxTotal = vData
End Property

Public Property Get TaxTotal() As Single
    TaxTotal = mvarTaxTotal
End Property
Public Property Let TktTaxTotal(ByVal vData As Single)
    mvarTktTaxTotal = vData
End Property

Public Property Get TktTaxTotal() As Single
    TktTaxTotal = mvarTktTaxTotal
End Property

Public Property Let BaseAmount(ByVal vData As Single)
    mvarBaseAmount = vData
End Property

Public Property Get BaseAmount() As Single
    BaseAmount = mvarBaseAmount
End Property
Public Property Let TktBaseAmount(ByVal vData As Single)
    mvarTktBaseAmount = vData
End Property

Public Property Get TktBaseAmount() As Single
    TktBaseAmount = mvarTktBaseAmount
End Property
Public Property Let Commission(ByVal vData As Single)
    mvarCommission = vData
End Property

Public Property Get Commission() As Single
    Commission = mvarCommission
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

Public Property Let SegmentSelectString(ByVal vData As String)
     mvarSegmentString = vData
End Property

Public Property Get SegmentSelectString() As String
     SegmentSelectString = mvarSegmentString
End Property

Public Function Tax(ByVal Index As Variant) As Tax
    Set Tax = mcolTaxes.Item(Index)
End Function
Public Function TktTax(ByVal Index As Variant) As Tax
    Set TktTax = mcolTktTaxes.Item(Index)
End Function

Public Sub TaxAdd(ByVal NewTax As Tax)
mcolTaxes.Add NewTax
'mvarTaxTotal = mvarTaxTotal + NewTax.Amount
'Zaw - V1.1.1 20100831 - IR4 - Extra decimal placing in Tax box in Fare Quote screen
mvarTaxTotal = CSng(CCur(CSng(mvarTaxTotal) + CSng(NewTax.Amount)))
End Sub
Public Sub TktTaxAdd(ByVal NewTax As Tax)
mcolTktTaxes.Add NewTax
mvarTktTaxTotal = mvarTktTaxTotal + NewTax.Amount
End Sub
Public Property Get TaxCount() As Long
    TaxCount = mcolTaxes.Count
End Property
Public Property Get TktTaxCount() As Long
    TktTaxCount = mcolTktTaxes.Count
End Property

Public Sub FareSegAdd(ByVal NewSegment As FareSegment)
mcolFareSegs.Add NewSegment
Set NewSegment = Nothing

End Sub
Public Function FareSeg(ByVal Index As Long) As FareSegment
    Set FareSeg = mcolFareSegs.Item(Index)
    
End Function

Public Sub SurchargeAdd(ByVal NewSurcharge As Surcharge)
    mcolSurcharges.Add NewSurcharge
    Set NewSurcharge = Nothing
End Sub

Public Function Surcharge(ByVal Index As Long) As Surcharge
    Set Surcharge = mcolSurcharges.Item(Index)
End Function

Public Property Get SurchargeCount() As Long
    SurchargeCount = mcolSurcharges.Count
End Property

Public Property Get FareSegCount() As Long
    FareSegCount = mcolFareSegs.Count
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
Public Property Let TktTotAmount(ByVal vData As Single)
    mvarTktTotAmount = vData
End Property

Public Property Get TktTotAmount() As Single
    TktTotAmount = mvarTktTotAmount
End Property

Public Property Let TotalCurrency(ByVal vData As String)
    mvarTotalCurrency = vData
End Property

Public Property Get TotalCurrency() As String
    TotalCurrency = mvarTotalCurrency
End Property
Public Property Let TktTotalCurrency(ByVal vData As String)
    mvarTktTotalCurrency = vData
End Property

Public Property Get TktTotalCurrency() As String
    TktTotalCurrency = mvarTktTotalCurrency
End Property

Public Property Let PlatCarrier(ByVal vData As String)
    mvarPlatCarrier = vData
End Property

Public Property Get PlatCarrier() As String
    PlatCarrier = mvarPlatCarrier
End Property

Private Sub Class_Initialize()
    Set mcolTaxes = New Collection
    Set mcolTktTaxes = New Collection
    Set mcolSurcharges = New Collection
    Set mcolFareSegs = New Collection
    Set mcolInfoMsg = New Collection
    Set mcolFareComponents = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolTaxes = Nothing
    Set mcolTktTaxes = Nothing
    Set mcolSurcharges = Nothing
    Set mcolFareSegs = Nothing
    Set mcolInfoMsg = Nothing
    Set mcolFareComponents = Nothing
    Set gobjLog = Nothing
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

Public Property Let CommissionPt(ByVal vData As Single)
    mvarCommissionPt = vData
End Property

Public Property Get CommissionPt() As Single
    CommissionPt = mvarCommissionPt
End Property

Public Property Let PTkt(ByVal vData As Boolean)
    mvarPTkt = vData
End Property

Public Property Get PTkt() As Boolean
    PTkt = mvarPTkt
End Property

Public Property Let RuleID(ByVal vData As String)
   mvarRuleID = vData
End Property

Public Property Get RuleID() As String
    RuleID = mvarRuleID
End Property

'Added by JiYong - Add in Actual Net Fare field
Public Property Let ActualNetAmount(ByVal vData As Single)
    mvarActualNetAmount = vData
End Property

Public Property Get ActualNetAmount() As Single
    ActualNetAmount = mvarActualNetAmount
End Property
