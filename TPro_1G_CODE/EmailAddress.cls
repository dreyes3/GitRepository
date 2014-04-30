VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmailAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarEItin As Boolean
Private mvarETkt As Boolean
Private mvarEInv As Boolean
Private mvarPaxType As String
Private mvarEmailAddress As String
Private mvarPNRLoc As String
Private mvarLineNum As Integer

Public Property Let EItin(ByVal vData As Boolean)
    mvarEItin = vData
End Property

Public Property Get EItin() As Boolean
    EItin = mvarEItin
End Property

Public Property Let ETkt(ByVal vData As Boolean)
    mvarETkt = vData
End Property

Public Property Get ETkt() As Boolean
    ETkt = mvarETkt
End Property

Public Property Let EInv(ByVal vData As Boolean)
    mvarEInv = vData
End Property

Public Property Get EInv() As Boolean
    EInv = mvarEInv
End Property

Public Property Let PaxType(ByVal vData As String)
    mvarPaxType = vData
End Property

Public Property Get PaxType() As String
    PaxType = mvarPaxType
End Property

Public Property Let EmailAddress(ByVal vData As String)
    mvarEmailAddress = vData
End Property

Public Property Get EmailAddress() As String
    EmailAddress = mvarEmailAddress
End Property

Public Property Let PNRLoc(ByVal vData As String)
    mvarPNRLoc = vData
End Property

Public Property Get PNRLoc() As String
    PNRLoc = mvarPNRLoc
End Property

Public Property Let LineNum(ByVal vData As Integer)
    mvarLineNum = vData
End Property

Public Property Get LineNum() As Integer
    LineNum = mvarLineNum
End Property

