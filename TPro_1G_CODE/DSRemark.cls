VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
'Data Standardization Information

Private mintDSItemNum As Integer 'local copy
Private mstrDSQualifier As String 'local copy
Private mstrDSRemarkText As String 'local copy
Private mdblDSPhaseNum As Double 'local copy
Public Property Let RemarkText(ByVal RmkText As String)
    mstrDSRemarkText = RmkText
End Property

Public Property Get RemarkText() As String
    RemarkText = mstrDSRemarkText
End Property

Public Property Let Qualifier(ByVal QualifierCode As String)
    mstrDSQualifier = QualifierCode
End Property

Public Property Get Qualifier() As String
    Qualifier = mstrDSQualifier
End Property

Public Property Let ItemNum(ByVal ItemNumber As Integer)
    mintDSItemNum = ItemNumber
End Property

Public Property Get ItemNum() As Integer
    ItemNum = mintDSItemNum
End Property

Public Property Let DSPhaseNum(ByVal PhaseNum As Double)
    mdblDSPhaseNum = PhaseNum
End Property

Public Property Get DSPhaseNum() As Double
    DSPhaseNum = mdblDSPhaseNum
End Property

