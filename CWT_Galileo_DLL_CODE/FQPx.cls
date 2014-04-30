VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcolFQ As Collection
'Private mvarFQ As FareQuote

'Public Property Let FQ(ByVal vData As FareQuote)
'    mvarFQ = vData
'End Property

'Public Property Get FQ() As FareQuote
'    FQ = mvarFQ
'End Property


Public Function FQ(ByVal Index As Variant) As FareQuote
    Set FQ = mcolFQ.Item(Index)
End Function

Public Sub FQAdd(ByVal NewFQ As FareQuote, Optional intBefore As Integer)
   If intBefore = 1 Then
      mcolFQ.Add NewFQ, , 1
   Else
      mcolFQ.Add NewFQ
   End If
End Sub



Public Property Get FQCount() As Long
    FQCount = mcolFQ.Count
End Property

Private Sub Class_Initialize()
    Set mcolFQ = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolFQ = Nothing
End Sub

