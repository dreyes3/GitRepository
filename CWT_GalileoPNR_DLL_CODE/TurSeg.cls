VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TurSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'TUR
Option Explicit

Private mvarSegNum As Integer 'local copy
Private mvarStatus As String
Private mvarSegType As String 'local copy
Private mvarStartDt As Date 'local copy
Private mvarEndDt As Date
Private mvarVnd As String
Private mvarNumPersons As Integer  'local copy
Private mvarNumNight As Integer 'local copy
Private mvarStartPt As String 'local copy
Private mvarEndPt As String 'local copy
Private mvarSellType As String 'local copy
Private mvarText As String 'local copy

Public Property Let SegNum(ByVal vData As Integer)
    mvarSegNum = vData
End Property

Public Property Get SegNum() As Integer
    SegNum = mvarSegNum
End Property

Public Property Let Status(ByVal vData As String)
    mvarStatus = vData
End Property

Public Property Get Status() As String
    SegType = mvarStatus
End Property

Public Property Let SegType(ByVal vData As String)
    mvarSegType = vData
End Property

Public Property Get SegType() As String
    SegType = mvarSegType
End Property

Public Property Let StartDt(ByVal vData As Date)
    mvarStartDt = vData
End Property

Public Property Get StartDt() As Date
    StartDt = mvarStartDt
End Property

Public Property Let EndDt(ByVal vData As Date)
    mvarEndDt = vData
End Property

Public Property Get EndDt() As Date
    EndDt = mvarEndDt
End Property

Public Property Let Vnd(ByVal vData As String)
    mvarVnd = vData
End Property

Public Property Get Vnd() As String
    Vnd = mvarSegType
End Property

Public Property Let NumPersons(ByVal vData As Integer)
    mvarNumPersons = vData
End Property

Public Property Get NumPersons() As Integer
    NumPersons = mvarNumPersons
End Property

Public Property Let NumNight(ByVal vData As Integer)
    mvarNumNight = vData
End Property

Public Property Get NumNight() As Integer
    NumNight = mvarNumNight
End Property

Public Property Let StartPt(ByVal vData As String)
    mvarStartPt = vData
End Property

Public Property Get StartPt() As String
    StartPt = mvarStartPt
End Property

Public Property Let EndPt(ByVal vData As String)
    mvarEndPt = vData
End Property

Public Property Get EndPt() As String
    EndPt = mvarEndPt
End Property

Public Property Let SellType(ByVal vData As String)
    mvarSellType = vData
End Property

Public Property Get SellType() As String
    SellType = mvarSellType
End Property

Public Property Let Text(ByVal vData As String)
    mvarText = vData
End Property


Public Property Get Text() As String
    Text = mvarText
End Property







