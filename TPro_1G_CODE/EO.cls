VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarEONumber As String
Private mvarProductCode As String
Private mvarVisaCountry As String
Private mvarVisaEntries As String
Private mvarVisaProcess As String
Private mvarVisaType As String
Private mvarContactNum As String
Private mvarProductSortKey As String
Private mvarVendorCode As String
Private mvarVendorName As String
Private mvarPNRRecLoc As String
Private mvarCN As String
Private mvarPaxName As String
Private mvarServiceDate As Date
Private mvarCost As Single
Private mvarSellPrice As Single
Private mvarCommissionAmt As Single
Private mvarFOP As String
Private mvarActionReq As Integer
Private mvarBillingDescription As String
Private mvarCreateDtTm As Date
Private mvarCreatedBy As String
Private mvarCreatedByName As String
Private mvarCreatedByPCC As String
Private mvarCompletedBy As String
Private mvarCompletedDtTm As Date
Private mvarEOType As String
Private mvarChequeNumber As String
Private mcolDescription As Collection
Private mcolRemarks As Collection
Private mcolTaxes As Collection
Private mvarDesctiptionLine As String
Private mvarTicketNumber As String
Private mvarTicketPrefix As String
Private mvarConjunctTicket As String
Private mstrFaxNo As String
Private mstrEmail As String
Private mstrReplyEmail As String
Private mstrContactPerson As String
Private mcolRIRemarks As Collection
Private mvarClientType As String
Private mvarNettFare As Double
Private mvarPublishedFare As Double
Private mvarGrossFare As Double
Private mvarDiscount As Double
Private mvarMerchFee As Double
Private mvarCWTAbsorb As Boolean
Private mvarTranxFee As Double
Private mvarSegSelect As String
'Private mvarTktNum As String
Private mvarListBoxRem As String
Private mvarPickUpFrom As String
Private mvarPickUpTo As String
Private mvarPickUpTime As Date
Private mvarPickUpFlight As String
Private mvarReturnFrom As String
Private mvarReturnTo As String
Private mvarReturnTime As Date
Private mvarReturnFlight As String
Private mvarAdditionalInfo As String
Private mvarVisaInfo As String
Private mvarAddress1 As String
Private mvarAddress2 As String
Private mvarCity As String
Private mvarCountry As String
Private mvarMISC As Boolean
Private mvarEOTaxTotal As Double
Private mstrRF As String
Private mstrLF As String
'CS Change EC
'Private mstrEC As String
Private mstrRS As String
Private mstrMS As String
Private mstrFF7 As String
Private mstrFF8 As String
'CS Remove FF26
'Private mstrFF26 As String
'CS Add FF41
Private mstrFF41 As String
Private mstrFF10 As String
Private mstrFF11 As String
Private mstrFF19 As String
Private mstrFF38 As String
Private mstrPassengerID As String
Private mcurSF As Currency
Private mblnTktType As Boolean
Private mblnTFNRCC As Boolean
Private mvarNettGST As Double
Private mstrFF81 As String
Private mstrRelateSegment As String
Private mstrBookingTool As String
Private mstrBookingAction As String
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private mstrPlatingCarrier As String
Private mbolWebFareApplied As Boolean
'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
Private mintFareType As Integer
'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
Private mstrTktNumber As String
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
Private mvarVendorHandling As Double

Public Property Let BookingTool(ByVal vData As String)
    mstrBookingTool = vData
End Property
Public Property Get BookingTool() As String
     BookingTool = mstrBookingTool
End Property
Public Property Let BookingAction(ByVal vData As String)
    mstrBookingAction = vData
End Property
Public Property Get BookingAction() As String
     BookingAction = mstrBookingAction
End Property
Public Property Let TFNRCC(ByVal vData As Boolean)
    mblnTFNRCC = vData
End Property

Public Property Get TFNRCC() As Boolean
    TFNRCC = mblnTFNRCC
End Property
Public Property Let TktType(ByVal vData As Boolean)
    mblnTktType = vData
End Property

Public Property Get TktType() As Boolean
    TktType = mblnTktType
End Property

Public Property Let RelateSegment(ByVal vData As String)
    mstrRelateSegment = vData
End Property
Public Property Get RelateSegment() As String
     RelateSegment = mstrRelateSegment
End Property
Public Property Let PassengerID(ByVal vData As String)
    mstrPassengerID = vData
End Property
Public Property Get PassengerID() As String
     PassengerID = mstrPassengerID
End Property
Public Property Let RF(ByVal vData As String)
    mstrRF = vData
End Property

Public Property Get RF() As String
    RF = mstrRF
End Property

Public Property Let LF(ByVal vData As String)
    mstrLF = vData
End Property

Public Property Get LF() As String
    LF = mstrLF
End Property

'CS Change EC
'Public Property Let EC(ByVal vData As String)
'    mstrEC = vData
'End Property
'Public Property Get EC() As String
'    EC = mstrEC
'End Property

Public Property Let rs(ByVal vData As String)
    mstrRS = vData
End Property

Public Property Get rs() As String
    rs = mstrRS
End Property
Public Property Let MS(ByVal vData As String)
    mstrMS = vData
End Property

Public Property Get MS() As String
    MS = mstrMS
End Property


Public Property Let FF7(ByVal vData As String)
    mstrFF7 = vData
End Property

Public Property Get FF7() As String
    FF7 = mstrFF7
End Property

Public Property Let FF8(ByVal vData As String)
    mstrFF8 = vData
End Property

Public Property Get FF8() As String
    FF8 = mstrFF8
End Property

'CS Remove FF26
'Public Property Let FF26(ByVal vData As String)
'    mstrFF26 = vData
'End Property
'Public Property Get FF26() As String
'    FF26 = mstrFF26
'End Property

'CS Add FF41
Public Property Let FF41(ByVal vData As String)
    mstrFF41 = vData
End Property
Public Property Get FF41() As String
    FF41 = mstrFF41
End Property


Public Property Let FF10(ByVal vData As String)
    mstrFF10 = vData
End Property

Public Property Get FF10() As String
    FF10 = mstrFF10
End Property

Public Property Let FF11(ByVal vData As String)
    mstrFF11 = vData
End Property

Public Property Get FF11() As String
    FF11 = mstrFF11
End Property

Public Property Let FF19(ByVal vData As String)
    mstrFF19 = vData
End Property

Public Property Get FF19() As String
    FF19 = mstrFF19
End Property
Public Property Let FF38(ByVal vData As String)
    mstrFF38 = vData
End Property

Public Property Get FF38() As String
    FF38 = mstrFF38
End Property

Public Sub RIRemarkAdd(ByVal RIRemarkText As String)
    mcolRIRemarks.Add RIRemarkText
End Sub

Public Property Get RIRemarkCount() As Long
    RIRemarkCount = mcolRIRemarks.Count
End Property

Public Property Get RIRemark(Index As Long)
    RIRemark = mcolRIRemarks.item(Index)
End Property

Public Property Get DescriptionLine(Index As Long) As String
    DescriptionLine = mcolDescription.item(Index)
End Property

Public Sub DescriptionLineAdd(ByVal DescriptionText As String)
    mcolDescription.Add DescriptionText
End Sub

Public Property Get DescriptionLinesCount() As Long
    DescriptionLinesCount = mcolDescription.Count
End Property

Public Property Get RemarkCount() As Long
    RemarkCount = mcolRemarks.Count
End Property

Public Sub RemarkAdd(ByVal RemarkText As String)
    mcolRemarks.Add RemarkText
End Sub

Public Property Get Remark(Index As Long)
    Remark = mcolRemarks.item(Index)
End Property

Public Property Let ChequeNumber(ByVal vData As String)
    mvarChequeNumber = vData
End Property

Public Property Get ChequeNumber() As String
    ChequeNumber = mvarChequeNumber
End Property

Public Property Let EOType(ByVal vData As String)
    mvarEOType = vData
End Property

Public Property Get EOType() As String
    EOType = mvarEOType
End Property

Public Property Let CompletedDtTm(ByVal vData As Date)
    mvarCompletedDtTm = vData
End Property

Public Property Get CompletedDtTm() As Date
    CompletedDtTm = mvarCompletedDtTm
End Property

Public Property Let CompletedBy(ByVal vData As String)
    mvarCompletedBy = vData
End Property

Public Property Get CompletedBy() As String
    CompletedBy = mvarCompletedBy
End Property

Public Property Let CreatedByName(ByVal vData As String)
    mvarCreatedByName = vData
End Property

Public Property Get CreatedByName() As String
    CreatedByName = mvarCreatedByName
End Property

Public Property Let CreatedBy(ByVal vData As String)
    mvarCreatedBy = vData
End Property

Public Property Get CreatedBy() As String
    CreatedBy = mvarCreatedBy
End Property
Public Property Let CreatedByPCC(ByVal vData As String)
    mvarCreatedByPCC = vData
End Property

Public Property Get CreatedByPCC() As String
    CreatedByPCC = mvarCreatedByPCC
End Property

Public Property Let CreateDtTm(ByVal vData As Date)
    mvarCreateDtTm = vData
End Property

Public Property Get CreateDtTm() As Date
    CreateDtTm = mvarCreateDtTm
End Property

Public Property Let BillingDescription(ByVal vData As String)
    mvarBillingDescription = vData
End Property

Public Property Get BillingDescription() As String
    BillingDescription = mvarBillingDescription
End Property

Public Property Let ActionReq(ByVal vData As Integer)
    mvarActionReq = vData
End Property

Public Property Get ActionReq() As Integer
    ActionReq = mvarActionReq
End Property

Public Property Let FOP(ByVal vData As String)
    mvarFOP = vData
End Property

Public Property Get FOP() As String
    FOP = mvarFOP
End Property

Public Property Let CommissionAmt(ByVal vData As Single)
    mvarCommissionAmt = vData
End Property

Public Property Get CommissionAmt() As Single
    CommissionAmt = mvarCommissionAmt
End Property

Public Property Let SellPrice(ByVal vData As Single)
    mvarSellPrice = vData
End Property

Public Property Get SellPrice() As Single
    SellPrice = mvarSellPrice
End Property

Public Property Let Cost(ByVal vData As Single)
    mvarCost = vData
End Property

Public Property Get Cost() As Single
    Cost = mvarCost
End Property

Public Property Let ServiceDate(ByVal vData As Date)
    mvarServiceDate = vData
End Property

Public Property Get ServiceDate() As Date
    ServiceDate = mvarServiceDate
End Property

Public Property Let PaxName(ByVal vData As String)
    mvarPaxName = vData
End Property

Public Property Get PaxName() As String
    PaxName = mvarPaxName
End Property

Public Property Let CN(ByVal vData As String)
    mvarCN = vData
End Property

Public Property Get CN() As String
    CN = mvarCN
End Property

Public Property Let PNRRecLoc(ByVal vData As String)
    mvarPNRRecLoc = vData
End Property

Public Property Get PNRRecLoc() As String
    PNRRecLoc = mvarPNRRecLoc
End Property

Public Property Let VendorCode(ByVal vData As String)
    mvarVendorCode = vData
End Property

Public Property Get VendorCode() As String
    VendorCode = mvarVendorCode
End Property

Public Property Let VendorName(ByVal vData As String)
    mvarVendorName = vData
End Property

Public Property Get VendorName() As String
    VendorName = mvarVendorName
End Property

Public Property Let Email(ByVal vData As String)
    mstrEmail = vData
End Property

Public Property Get Email() As String
    Email = mstrEmail
End Property
Public Property Let ReplyEmail(ByVal vData As String)
    mstrReplyEmail = vData
End Property

Public Property Get ReplyEmail() As String
    ReplyEmail = mstrReplyEmail
End Property


Public Property Let ContactPerson(ByVal vData As String)
    mstrContactPerson = vData
End Property

Public Property Get ContactPerson() As String
    ContactPerson = mstrContactPerson
End Property

Public Property Let FaxNo(ByVal vData As String)
    mstrFaxNo = vData
End Property

Public Property Get FaxNo() As String
    FaxNo = mstrFaxNo
End Property

Public Property Let ProductCode(ByVal vData As String)
    mvarProductCode = vData
End Property

Public Property Get ProductCode() As String
    ProductCode = mvarProductCode
End Property

Public Property Let ContactNum(ByVal vData As String)
    mvarContactNum = vData
End Property

Public Property Get ContactNum() As String
    ContactNum = mvarContactNum
End Property

Public Property Let ProductSortKey(ByVal vData As String)
    mvarProductSortKey = vData
End Property

Public Property Get ProductSortKey() As String
    ProductSortKey = mvarProductSortKey
End Property

Public Property Let EONumber(ByVal vData As String)
    mvarEONumber = vData
End Property

Public Property Get EONumber() As String
    EONumber = mvarEONumber
End Property
Public Sub TaxAdd(ByVal TaxAmount As Single, ByVal TaxCode As String)
    Dim objTax As New EOTax
    
    objTax.Amount = TaxAmount
    objTax.Code = TaxCode
    
    mcolTaxes.Add objTax
    mvarEOTaxTotal = mvarEOTaxTotal + TaxAmount
    
End Sub
Public Property Get EOTaxTotal() As Double
    EOTaxTotal = mvarEOTaxTotal
End Property
Public Property Get Tax(Index As Long) As EOTax
    Set Tax = mcolTaxes.item(Index)
End Property

Public Property Get TaxCount() As Long
    TaxCount = mcolTaxes.Count
End Property

Public Property Let TicketNumber(ByVal vData As String)
    mvarTicketNumber = vData
End Property

Public Property Get TicketNumber() As String
    TicketNumber = mvarTicketNumber
End Property

Public Property Let ConjunctTicket(ByVal vData As String)
    mvarConjunctTicket = vData
End Property

Public Property Get ConjunctTicket() As String
    ConjunctTicket = mvarConjunctTicket
End Property

Public Property Let ClientType(ByVal vData As String)
    mvarClientType = vData
End Property

Public Property Get ClientType() As String
    ClientType = mvarClientType
End Property

Public Property Let NettFare(ByVal vData As Double)
    mvarNettFare = vData
End Property

Public Property Get NettFare() As Double
    NettFare = mvarNettFare
End Property

Public Property Let PublishedFare(ByVal vData As Double)
    mvarPublishedFare = vData
End Property

Public Property Get PublishedFare() As Double
    PublishedFare = mvarPublishedFare
End Property

Public Property Let GrossFare(ByVal vData As Double)
    mvarGrossFare = vData
End Property

Public Property Get GrossFare() As Double
    GrossFare = mvarGrossFare
End Property

Public Property Let Discount(ByVal vData As Double)
    mvarDiscount = vData
End Property

Public Property Get Discount() As Double
    Discount = mvarDiscount
End Property

Public Property Let MerchFee(ByVal vData As Double)
    mvarMerchFee = vData
End Property

Public Property Get MerchFee() As Double
    MerchFee = mvarMerchFee
End Property

Public Property Let CWTAbsorb(ByVal vData As Boolean)
    mvarCWTAbsorb = vData
End Property

Public Property Get CWTAbsorb() As Boolean
    CWTAbsorb = mvarCWTAbsorb
End Property

Public Property Let TranxFee(ByVal vData As Double)
    mvarTranxFee = vData
End Property

Public Property Get TranxFee() As Double
    TranxFee = mvarTranxFee
End Property

Public Property Let SegSelect(ByVal vData As String)
    mvarSegSelect = vData
End Property

Public Property Get SegSelect() As String
    SegSelect = mvarSegSelect
End Property

'Public Property Let TktNum(ByVal vData As String)
'    mvarTktNum = vData
'End Property
'
'Public Property Get TktNum() As String
'    TktNum = mvarTktNum
'End Property

Public Property Let ListBoxRem(ByVal vData As String)
    mvarListBoxRem = vData
End Property

Public Property Get ListBoxRem() As String
    ListBoxRem = mvarListBoxRem
End Property

Public Property Let PickUpFrom(ByVal vData As String)
    mvarPickUpFrom = vData
End Property

Public Property Get PickUpFrom() As String
    PickUpFrom = mvarPickUpFrom
End Property

Public Property Let PickUpTo(ByVal vData As String)
    mvarPickUpTo = vData
End Property

Public Property Get PickUpTo() As String
    PickUpTo = mvarPickUpTo
End Property

Public Property Let PickUpTime(ByVal vData As Date)
    mvarPickUpTime = vData
End Property

Public Property Get PickUpTime() As Date
    PickUpTime = mvarPickUpTime
End Property

Public Property Let PickUpFlight(ByVal vData As String)
    mvarPickUpFlight = vData
End Property

Public Property Get PickUpFlight() As String
    PickUpFlight = mvarPickUpFlight
End Property

Public Property Let ReturnFrom(ByVal vData As String)
    mvarReturnFrom = vData
End Property

Public Property Get ReturnFrom() As String
    ReturnFrom = mvarReturnFrom
End Property

Public Property Let ReturnTo(ByVal vData As String)
    mvarReturnTo = vData
End Property

Public Property Get ReturnTo() As String
    ReturnTo = mvarReturnTo
End Property

Public Property Let ReturnTime(ByVal vData As Date)
    mvarReturnTime = vData
End Property

Public Property Get ReturnTime() As Date
    ReturnTime = mvarReturnTime
End Property

Public Property Let ReturnFlight(ByVal vData As String)
    mvarReturnFlight = vData
End Property

Public Property Get ReturnFlight() As String
    ReturnFlight = mvarReturnFlight
End Property

Public Property Let AdditionalInfo(ByVal vData As String)
    mvarAdditionalInfo = vData
End Property

Public Property Get AdditionalInfo() As String
    AdditionalInfo = mvarAdditionalInfo
End Property

Public Property Let VisaInfo(ByVal vData As String)
    mvarVisaInfo = vData
End Property

Public Property Get VisaInfo() As String
    VisaInfo = mvarVisaInfo
End Property

Public Property Let Address1(ByVal vData As String)
    mvarAddress1 = vData
End Property

Public Property Get Address1() As String
    Address1 = mvarAddress1
End Property

Public Property Let Address2(ByVal vData As String)
    mvarAddress2 = vData
End Property

Public Property Get Address2() As String
    Address2 = mvarAddress2
End Property

Public Property Let City(ByVal vData As String)
    mvarCity = vData
End Property

Public Property Get City() As String
    City = mvarCity
End Property

Public Property Let Country(ByVal vData As String)
    mvarCountry = vData
End Property

Public Property Get Country() As String
    Country = mvarCountry
End Property
Public Property Let VisaCountry(ByVal vData As String)
    mvarVisaCountry = vData
End Property

Public Property Get VisaCountry() As String
    VisaCountry = mvarVisaCountry
End Property
Public Property Let VisaEntries(ByVal vData As String)
    mvarVisaEntries = vData
End Property

Public Property Get VisaEntries() As String
    VisaEntries = mvarVisaEntries
End Property
Public Property Let VisaProcess(ByVal vData As String)
    mvarVisaProcess = vData
End Property

Public Property Get VisaProcess() As String
    VisaProcess = mvarVisaProcess
End Property
Public Property Let Misc(ByVal vData As Boolean)
    mvarMISC = vData
End Property

Public Property Get Misc() As Boolean
    Misc = mvarMISC
End Property

Public Property Let SF(ByVal vData As Currency)
    mcurSF = vData
End Property

Public Property Get SF() As Currency
    SF = mcurSF
End Property
Public Property Let NettGST(ByVal vData As Double)
    mvarNettGST = vData
End Property

Public Property Get NettGST() As Double
    NettGST = mvarNettGST
End Property

Private Sub Class_Initialize()
    Set mcolDescription = New Collection
    Set mcolRemarks = New Collection
    Set mcolTaxes = New Collection
    Set mcolRIRemarks = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolDescription = Nothing
    Set mcolRemarks = Nothing
    Set mcolTaxes = Nothing
    Set mcolRIRemarks = Nothing
End Sub
Public Property Let VisaType(ByVal vData As String)
    mvarVisaType = vData
End Property

Public Property Get VisaType() As String
    VisaType = mvarVisaType
End Property

Public Property Let FF81(ByVal vData As String)
    mstrFF81 = vData
End Property

Public Property Get FF81() As String
    FF81 = mstrFF81
End Property

Public Property Let PlatingCarrier(ByVal vData As String)
    mstrPlatingCarrier = vData
End Property

Public Property Get PlatingCarrier() As String
    PlatingCarrier = mstrPlatingCarrier
End Property

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Public Property Let WebFareApplied(ByVal vData As Boolean)
    mbolWebFareApplied = vData
End Property

Public Property Get WebFareApplied() As Boolean
    WebFareApplied = mbolWebFareApplied
End Property
'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
Public Property Let FareType(ByVal vData As Integer)
    mintFareType = vData
End Property

Public Property Get FareType() As Integer
    FareType = mintFareType
End Property

'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
Public Property Let TktNumber(ByVal vData As String)
    mstrTktNumber = vData
End Property
Public Property Get TktNumber() As String
     TktNumber = mstrTktNumber
End Property
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
Public Property Let VendorHandling(ByVal vData As Double)
    mvarVendorHandling = vData
End Property

'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
Public Property Get VendorHandling() As Double
    VendorHandling = mvarVendorHandling
End Property
