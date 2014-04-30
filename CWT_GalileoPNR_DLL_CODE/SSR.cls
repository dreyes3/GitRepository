VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mintGFax As Integer
Private mstrSSCode As String
Private mstrStatus As String
Private mintSegNum As Integer
Private mintPsgNum As Integer
Private mstrText As String
Public Property Let GFax(ByVal vData As Integer)
    mintGFax = vData
End Property
Public Property Get GFax() As Integer
    GFax = mintGFax
End Property
Public Property Let SSCode(ByVal vData As String)
    mstrSSCode = vData
End Property
Public Property Get SSCode() As String
    SSCode = mstrSSCode
End Property
Public Property Let Status(ByVal vData As String)
    mstrStatus = vData
End Property
Public Property Get Status() As String
    Status = mstrStatus
End Property
Public Property Let SegNum(ByVal vData As Integer)
    mintSegNum = vData
End Property
Public Property Get SegNum() As Integer
    SegNum = mintSegNum
End Property
Public Property Let PsgNum(ByVal vData As Integer)
    mintPsgNum = vData
End Property
Public Property Get PsgNum() As Integer
    PsgNum = mintPsgNum
End Property
Public Property Let Text(ByVal vData As String)
    mstrText = vData
End Property
Public Property Get Text() As String
    Text = mstrText
End Property



