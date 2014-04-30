VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarWONum As String 'local copy
Private mvarCompanyName As String 'local copy
Private mvarProfileName As String 'local copy
Private mvarClientType As String 'local copy
Private mvarDIV As String 'local copy
Private mvarCreditTerms As Long 'local copy
Private mvarCreditStatus As Byte 'local copy
Private mvarMerchFeePct As Single
Private mvarDiscountDomestic As Single 'local copy
Private mvarDiscountInternational As Single 'local copy
Private mvarMarkUp As Single 'local copy
Private mvarTransactionFeeGroup As Integer 'local copy
Private mvarMarkUpGroup As Integer 'local copy
Private mvarHotelRateCode As String 'local copy
Private mvarContracts As Boolean 'local copy
Private mvarEInvoice As Boolean 'local copy
Private mvarClientID As Integer 'local copy
Private mvarTFIncMF As Boolean
Private mvarCPG As Boolean
Private mvarAltFQPCC As String
Private mvarTFCalcBy As String
Private mvarMI As Boolean
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Insert new value to store the Profile PCC
Private mvarProfilePCC As String
'Preethi - V1.2.4 20110622  - ER01 - Tracking of Touches to CWT Booking
Private mvarPNRTrackingTouches As Boolean
'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
Private mvarAquaItin As Boolean
'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
Private mstrAgencyName As String
'--

Public Property Let MI(ByVal vData As Boolean)
    mvarMI = vData
End Property

Public Property Get MI() As Boolean
    MI = mvarMI
End Property

Public Property Let AltFQPCC(ByVal vData As String)
    mvarAltFQPCC = vData
End Property

Public Property Get AltFQPCC() As String
    AltFQPCC = mvarAltFQPCC
End Property
Public Property Let CPG(ByVal vData As Boolean)
    mvarCPG = vData
End Property

Public Property Get CPG() As Boolean
    CPG = mvarCPG
End Property
Public Property Let TFIncMF(ByVal vData As Boolean)
    mvarTFIncMF = vData
End Property

Public Property Get TFIncMF() As Boolean
    TFIncMF = mvarTFIncMF
End Property
Public Property Let EInvoice(ByVal vData As Boolean)
    mvarEInvoice = vData
End Property

Public Property Get EInvoice() As Boolean
    EInvoice = mvarEInvoice
End Property

Public Property Let Contracts(ByVal vData As Boolean)
    mvarContracts = vData
End Property

Public Property Get Contracts() As Boolean
    Contracts = mvarContracts
End Property

Public Property Let HotelRateCode(ByVal vData As String)
    mvarHotelRateCode = vData
End Property

Public Property Get HotelRateCode() As String
    HotelRateCode = mvarHotelRateCode
End Property

Public Property Let TransactionFeeGroup(ByVal vData As Integer)
    mvarTransactionFeeGroup = vData
End Property

Public Property Get TransactionFeeGroup() As Integer
    TransactionFeeGroup = mvarTransactionFeeGroup
End Property

Public Property Let MarkUpGroup(ByVal vData As Integer)
    mvarMarkUpGroup = vData
End Property

Public Property Get MarkUpGroup() As Integer
    MarkUpGroup = mvarMarkUpGroup
End Property

Public Property Let MarkUp(ByVal vData As Single)
    mvarMarkUp = vData
End Property

Public Property Get MarkUp() As Single
    MarkUp = mvarMarkUp
End Property

Public Property Let DiscountInternational(ByVal vData As Single)
    mvarDiscountInternational = vData
End Property

Public Property Get DiscountInternational() As Single
    DiscountInternational = mvarDiscountInternational
End Property

Public Property Let DiscountDomestic(ByVal vData As Single)
    mvarDiscountDomestic = vData
End Property

Public Property Get DiscountDomestic() As Single
    DiscountDomestic = mvarDiscountDomestic
End Property

Public Property Let CreditStatus(ByVal vData As Byte)
    mvarCreditStatus = vData
End Property

Public Property Get CreditStatus() As Byte
    CreditStatus = mvarCreditStatus
End Property

Public Property Let CreditTerms(ByVal vData As Long)
    mvarCreditTerms = vData
End Property

Public Property Get CreditTerms() As Long
    CreditTerms = mvarCreditTerms
End Property

Public Property Let MerchFeePct(ByVal vData As Single)
    mvarMerchFeePct = vData
End Property

Public Property Get MerchFeePct() As Single
    MerchFeePct = mvarMerchFeePct
End Property

Public Property Let DIV(ByVal vData As String)
    mvarDIV = vData
End Property

Public Property Get DIV() As String
    DIV = mvarDIV
End Property

Public Property Let ClientType(ByVal vData As String)
    mvarClientType = vData
End Property

Public Property Get ClientType() As String
    ClientType = mvarClientType
End Property

Public Property Let ProfileName(ByVal vData As String)
    mvarProfileName = vData
End Property

Public Property Get ProfileName() As String
    ProfileName = mvarProfileName
End Property

Public Property Let CompanyName(ByVal vData As String)
    mvarCompanyName = vData
End Property

Public Property Get CompanyName() As String
    CompanyName = mvarCompanyName
End Property

Public Property Let WONum(ByVal vData As String)
    mvarWONum = vData
End Property

Public Property Get WONum() As String
    WONum = mvarWONum
End Property

Public Property Let ClientID(ByVal vData As Integer)
    mvarClientID = vData
End Property

Public Property Get ClientID() As Integer
    ClientID = mvarClientID
End Property


Public Property Let TFCalcBy(ByVal vData As String)
    mvarTFCalcBy = vData
End Property

Public Property Get TFCalcBy() As String
    TFCalcBy = mvarTFCalcBy
End Property

'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Insert new value to store the Profile PCC
Public Property Let ProfilePCC(ByVal vData As String)
    mvarProfilePCC = vData
End Property

Public Property Get ProfilePCC() As String
    ProfilePCC = mvarProfilePCC
End Property
'Preethi - V1.2.4 20110622  - ER01 - Tracking of Touches to CWT Booking
Public Property Let PNRTrackingTouches(ByVal vData As Boolean)
    mvarPNRTrackingTouches = vData
End Property

Public Property Get PNRTrackingTouches() As Boolean
    PNRTrackingTouches = mvarPNRTrackingTouches
End Property
'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
Public Property Let AquaItin(ByVal vData As Boolean)
    mvarAquaItin = vData
End Property

Public Property Get AquaItin() As Boolean
    AquaItin = mvarAquaItin
End Property

'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
Public Property Let AgencyName(ByVal vData As String)
    mstrAgencyName = vData
End Property
'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
Public Property Get AgencyName() As String
    AgencyName = mstrAgencyName
End Property
