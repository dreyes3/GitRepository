VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareOptionSegments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mcolFareOptionSegments As Collection

Public Sub AddFareOptionSegment(ByVal FareOptionSegment As FareOptionSegment)
    mcolFareOptionSegments.Add FareOptionSegment
End Sub

Public Property Get FareOptionSegment(Index As Integer) As FareOptionSegment
    Set FareOptionSegment = mcolFareOptionSegments.item(Index)
End Property

Public Property Get FareOptionSegmentCount() As Integer
    FareOptionSegmentCount = mcolFareOptionSegments.Count
End Property

Private Sub Class_Initialize()
    Set mcolFareOptionSegments = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolFareOptionSegments = Nothing
End Sub



