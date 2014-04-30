VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mintItemNum As Integer 'local copy
Private mstrQualifier As String 'local copy
Private mstrRemarkText As String 'local copy
Public Property Let RemarkText(ByVal RmkText As String)
    mstrRemarkText = RmkText
End Property

Public Property Get RemarkText() As String
    RemarkText = mstrRemarkText
End Property

Public Property Let Qualifier(ByVal QualifierCode As String)
    mstrQualifier = QualifierCode
End Property

Public Property Get Qualifier() As String
    Qualifier = mstrQualifier
End Property

Public Property Let ItemNum(ByVal ItemNumber As Integer)
    mintItemNum = ItemNumber
End Property

Public Property Get ItemNum() As Integer
    ItemNum = mintItemNum
End Property

