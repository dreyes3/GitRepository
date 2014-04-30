VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebFares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration

Private mcolWebFares As Collection

Public Sub AddWebFare(ByVal WebFare As WebFare)
    mcolWebFares.Add WebFare
End Sub

Public Property Get WebFare(Index As Integer) As WebFare
    Set WebFare = mcolWebFares.item(Index)
End Property

Public Property Get WebFareCount() As Integer
    WebFareCount = mcolWebFares.Count
End Property

Private Sub Class_Initialize()
    Set mcolWebFares = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolWebFares = Nothing
End Sub


