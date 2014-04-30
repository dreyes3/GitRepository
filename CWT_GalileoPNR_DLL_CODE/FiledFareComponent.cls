VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FiledFareComponent"
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

Friend Property Let Amount(ByVal FCAmount As Single)
    mdblAmount = FCAmount
End Property

Public Property Get Amount() As Single
    Amount = mdblAmount
End Property

Friend Property Let Net(ByVal NetFlag As Boolean)
    mbolNet = NetFlag
End Property

Public Property Get Net() As Boolean
    Net = mbolNet
End Property

Friend Property Let RuleNum(ByVal Code As String)
    mstrRuleNum = Code
End Property

Public Property Get RuleNum() As String
    RuleNum = mstrRuleNum
End Property

Friend Property Let CurrencyCode(ByVal Code As String)
    mstrCurrencyCode = Code
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = mstrCurrencyCode
End Property

Friend Property Let Destinantion(ByVal DestCityCode As String)
    mstrDestinantion = DestCityCode
End Property

Public Property Get Destinantion() As String
    Destinantion = mstrDestinantion
End Property

Friend Property Let Origin(ByVal OrigCityCode As String)
    mstrOrigin = OrigCityCode
End Property

Public Property Get Origin() As String
    Origin = mstrOrigin
End Property

Friend Property Let FBC(ByVal FareCode As String)
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

