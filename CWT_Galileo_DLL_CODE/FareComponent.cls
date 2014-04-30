VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mintNumber As Integer 'local copy
Private mstrFBC As String 'local copy
Private mstrOrigin As String 'local copy
Private mstrDestinantion As String 'local copy
Private mstrCurrencyCode As String 'local copy
Private mstrRuleNum As String 'local copy
Private mbolNet As Boolean 'local copy
Private mdblAmount As Single 'local copy
Private mvarVendor As String 'local copy
Private mvarOTWCarrier As String 'local copy
Private mvarTPM As Long 'local copy
Private mvarMPM As Long 'local copy
Private mvarDirectionalInd As String
Private mvarPriceFBC As String 'local copy
Private mvarTicketFBC As String 'local copy
Private mvarValueCode As String 'local copy
Private mvarTktDesig As String
Private mvarCommissionOnTicket As Long 'local copy
Private mvarNRCC As Boolean 'local copy
Private mvarFOPCode As String 'local copy
Private mvarHPFApplies As Boolean 'local copy
Private mvarWLAllow As Boolean 'local copy
Private mvarOpenSegAllow As Boolean 'local copy
Private mvarETRequired As Boolean 'local copy
Private mvarPaperTktSurcharge As Single 'local copy
Private mvarManualProcRequired As Boolean 'local copy
Private mvarFareOnTicket As String 'local copy
Private mbolTktInfoApplies As Boolean
Private mcolEndorsements As Collection
Private mcolRules As Collection
Private mvarDestinationCountry As String 'local copy
Private mvarDestinationRegion As String 'local copy
Private mvarOriginCountry As String 'local copy
Private mvarOriginRegion As String 'local copy
Private mvarTourCode As String

Public Property Let OriginRegion(ByVal vData As String)
    mvarOriginRegion = vData
End Property

Public Property Get OriginRegion() As String
    OriginRegion = mvarOriginRegion
End Property

Public Property Let OriginCountry(ByVal vData As String)
    mvarOriginCountry = vData
End Property

Public Property Get OriginCountry() As String
    OriginCountry = mvarOriginCountry
End Property

Public Property Let DestinationRegion(ByVal vData As String)
    mvarDestinationRegion = vData
End Property

Public Property Get DestinationRegion() As String
    DestinationRegion = mvarDestinationRegion
End Property

Public Property Let DestinationCountry(ByVal vData As String)
    mvarDestinationCountry = vData
End Property

Public Property Get DestinationCountry() As String
    DestinationCountry = mvarDestinationCountry
End Property

Public Property Let FareOnTicket(ByVal vData As String)
    mvarFareOnTicket = vData
End Property

Public Property Get FareOnTicket() As String
    FareOnTicket = mvarFareOnTicket
End Property

Public Property Let TourCode(ByVal vData As String)
    mvarTourCode = vData
End Property

Public Property Get TourCode() As String
    TourCode = mvarTourCode
End Property

Public Property Let ManualProcRequired(ByVal vData As Boolean)
    mvarManualProcRequired = vData
End Property

Public Property Get ManualProcRequired() As Boolean
    ManualProcRequired = mvarManualProcRequired
End Property

Public Property Let PaperTktSurcharge(ByVal vData As Single)
    mvarPaperTktSurcharge = vData
End Property

Public Property Get PaperTktSurcharge() As Single
    PaperTktSurcharge = mvarPaperTktSurcharge
End Property

Public Property Let ETRequired(ByVal vData As Boolean)
    mvarETRequired = vData
End Property

Public Property Get ETRequired() As Boolean
    ETRequired = mvarETRequired
End Property

Public Property Let OpenSegAllow(ByVal vData As Boolean)
    mvarOpenSegAllow = vData
End Property

Public Property Get OpenSegAllow() As Boolean
    OpenSegAllow = mvarOpenSegAllow
End Property

Public Property Let WLAllow(ByVal vData As Boolean)
    mvarWLAllow = vData
End Property

Public Property Get WLAllow() As Boolean
    WLAllow = mvarWLAllow
End Property

Public Property Let HPFApplies(ByVal vData As Boolean)
    mvarHPFApplies = vData
End Property

Public Property Get HPFApplies() As Boolean
    HPFApplies = mvarHPFApplies
End Property

Public Property Let FOPCode(ByVal vData As String)
    mvarFOPCode = vData
End Property

Public Property Get FOPCode() As String
    FOPCode = mvarFOPCode
End Property

Public Property Let NRCC(ByVal vData As Boolean)
    mvarNRCC = vData
End Property

Public Property Get NRCC() As Boolean
    NRCC = mvarNRCC
End Property

Public Property Let CommissionOnTicket(ByVal vData As Long)
    mvarCommissionOnTicket = vData
End Property

Public Property Get CommissionOnTicket() As Long
    CommissionOnTicket = mvarCommissionOnTicket
End Property

Public Property Let ValueCode(ByVal vData As String)
    mvarValueCode = vData
End Property

Public Property Get ValueCode() As String
    ValueCode = mvarValueCode
End Property

Public Property Let TktDesignator(ByVal vData As String)
    mvarTktDesig = vData
End Property

Public Property Get TktDesignator() As String
    TktDesignator = mvarTktDesig
End Property

Public Property Let TicketFBC(ByVal vData As String)
    mvarTicketFBC = vData
End Property

Public Property Get TicketFBC() As String
    TicketFBC = mvarTicketFBC
End Property

Public Property Let PriceFBC(ByVal vData As String)
    mvarPriceFBC = vData
End Property

Public Property Get PriceFBC() As String
    PriceFBC = mvarPriceFBC
End Property

Public Property Let MPM(ByVal vData As Long)
    mvarMPM = vData
End Property

Public Property Get MPM() As Long
    MPM = mvarMPM
End Property

Public Property Let TPM(ByVal vData As Long)
    mvarTPM = vData
End Property

Public Property Get TPM() As Long
    TPM = mvarTPM
End Property

Public Property Let OTWCarrier(ByVal vData As String)
    mvarOTWCarrier = vData
End Property

Public Property Get OTWCarrier() As String
    OTWCarrier = mvarOTWCarrier
End Property

Public Property Let DirectionalInd(ByVal vData As String)
    mvarDirectionalInd = vData
End Property

Public Property Get DirectionalInd() As String
    DirectionalInd = mvarDirectionalInd
End Property

Public Property Let Vendor(ByVal vData As String)
    mvarVendor = vData
End Property

Public Property Get Vendor() As String
    Vendor = mvarVendor
End Property

Public Property Let Amount(ByVal FCAmount As Single)
    mdblAmount = FCAmount
End Property

Public Property Get Amount() As Single
    Amount = mdblAmount
End Property

Public Property Let Net(ByVal NetFlag As Boolean)
    mbolNet = NetFlag
End Property

Public Property Get Net() As Boolean
    Net = mbolNet
End Property

Public Property Let RuleNum(ByVal Code As String)
    mstrRuleNum = Code
End Property

Public Property Get RuleNum() As String
    RuleNum = mstrRuleNum
End Property

Public Property Let CurrencyCode(ByVal Code As String)
    mstrCurrencyCode = Code
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = mstrCurrencyCode
End Property

Public Property Let Destinantion(ByVal DestCityCode As String)
    mstrDestinantion = DestCityCode
End Property

Public Property Get Destinantion() As String
    Destinantion = mstrDestinantion
End Property

Public Property Let Origin(ByVal OrigCityCode As String)
    mstrOrigin = OrigCityCode
End Property

Public Property Get Origin() As String
    Origin = mstrOrigin
End Property

Public Property Let FBC(ByVal FareCode As String)
    mstrFBC = FareCode
End Property

Public Property Get FBC() As String
    FBC = mstrFBC
End Property

Friend Property Let Number(ByVal FCNumber As Integer)
    mintNumber = FCNumber
End Property

Public Property Get Number() As Integer
    Number = mintNumber
End Property

Private Sub Class_Initialize()
    Set mcolEndorsements = New Collection
    Set mcolRules = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set mcolEndorsements = Nothing
    Set mcolRules = Nothing
End Sub

Public Sub RuleTextAdd(ByVal FareRuleText As String)
    mcolRules.Add FareRuleText
End Sub

Public Property Get RuleTextCount() As Long
    RuleTextCount = mcolRules.Count
End Property

Public Sub EndorsementAdd(ByVal NewEndorsment As String, Optional EndorsementNumber As Long = 6)
    If EndorsementNumber <= mcolEndorsements.Count Then
        mcolEndorsements.Add NewEndorsment, , , EndorsementNumber
        mcolEndorsements.Remove EndorsementNumber
        
    ElseIf mcolEndorsements.Count < 6 Then
        mcolEndorsements.Add NewEndorsment
    End If
End Sub

Public Property Get EndorsementCount() As Long
    EndorsementCount = mcolEndorsements.Count
End Property

Public Property Let TktInfoApplies(ByVal vData As Boolean)
    mbolTktInfoApplies = vData
End Property

Public Property Get TktInfoApplies() As Boolean
    TktInfoApplies = mbolTktInfoApplies
End Property

Public Function Endorsement(ByVal Index As Long) As String
    If Index > mcolEndorsements.Count Then
        Endorsement = ""
    Else
        Endorsement = mcolEndorsements(Index)
    End If
End Function

Public Function RuleText(ByVal Index As Long) As String
    If Index > mcolRules.Count Then
        RuleText = ""
    Else
        RuleText = mcolRules(Index)
    End If
End Function

