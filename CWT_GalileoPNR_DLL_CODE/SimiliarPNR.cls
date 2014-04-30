VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimiliarName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarPaxName As String
Private mvarDate As Date
Private mvarSeat As Integer
Private mvarPNR As String
Private mvarCancel As Boolean

Public Property Let PaxName(ByVal vData As String)
    mvarPaxName = vData
End Property

Public Property Get PaxName() As String
    PaxName = mvarPaxName
End Property

Public Property Let FirstSegDate(ByVal vData As String)
    mvarDate = vData
End Property

Public Property Get FirstSegDate() As String
    FirstSegDate = mvarDate
End Property

Public Property Let Seat(ByVal vData As String)
    mvarSeat = vData
End Property

Public Property Get Seat() As String
    Seat = mvarSeat
End Property

Public Property Let PNR(ByVal vData As String)
    mvarPNR = vData
End Property

Public Property Get PNR() As String
    PNR = mvarPNR
End Property

Public Property Let Status(ByVal vData As String)
    mvarCancel = vData
End Property

Public Property Get Status() As String
    Status = mvarCancel
End Property
