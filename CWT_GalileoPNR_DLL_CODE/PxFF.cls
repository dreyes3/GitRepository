VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PxFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcolPX As Collection

Public Function PX(ByVal Index As Variant) As FiledFare
    Set PX = mcolPX.Item(Index)
End Function

Public Sub PXAdd(ByVal NewPX As FiledFare, Optional intBefore As Integer)
   If intBefore = 1 Then
      mcolPX.Add NewPX, , 1
   Else
      mcolPX.Add NewPX
   End If
End Sub

Public Property Get PXCount() As Long
    PXCount = mcolPX.Count
End Property

Private Sub Class_Initialize()
    Set mcolPX = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolPX = Nothing
End Sub
