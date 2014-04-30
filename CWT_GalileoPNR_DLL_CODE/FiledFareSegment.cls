VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FiledFareSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrFBC As String 'local copy
Private mdtmNVB As Date 'local copy
Private mdtmNVA As Date 'local copy
Private mstrTD As String 'local copy
Private mstrBagInfo As String 'local copy
Private mbolStopover As Boolean 'local copy
Private mstrInfoText As String 'local copy
Private mvarAssocSeg As String

Public Property Let AssocSeg(ByVal AssociatedSegment As String)
    mvarAssocSeg = AssociatedSegment
End Property

Public Property Get AssocSeg() As String
    AssocSeg = mvarAssocSeg
End Property

Public Property Let Stopover(ByVal StopInd As Boolean)
    mbolStopover = StopInd
End Property

Public Property Get Stopover() As Boolean
    Stopover = mbolStopover
End Property

Public Property Let BagInfo(ByVal BagData As String)
    mstrBagInfo = BagData
End Property

Public Property Get BagInfo() As String
    BagInfo = mstrBagInfo
End Property

Public Property Let TD(ByVal TicketDesig As String)
    mstrTD = TicketDesig
End Property

Public Property Get TD() As String
    TD = mstrTD
End Property

Public Property Let NVA(ByVal NotValidAfter As Date)
    mdtmNVA = NotValidAfter
End Property

Public Property Get NVA() As Date
    NVA = mdtmNVA
End Property

Public Property Let NVB(ByVal NotValidBefore As Date)
    mdtmNVB = NotValidBefore
End Property

Public Property Get NVB() As Date
    NVB = mdtmNVB
End Property

Public Property Let FBC(ByVal FareBasisCode As String)
    mstrFBC = FareBasisCode
End Property

Public Property Get FBC() As String
    FBC = mstrFBC
End Property

