VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Remarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 'Preethi - V1.2.8  20111031 - CR101 - Aqua Itin Remark Screen
Private mcolRemarks As Collection

Public Sub AddRemark(ByVal Remark As Remark)
    mcolRemarks.Add Remark
End Sub

Public Property Get Remark(Index As Long) As Remark
    Set Remark = mcolRemarks.item(Index)
End Property

Public Property Get RemarkCount() As Long
    RemarkCount = mcolRemarks.Count
End Property

Private Sub Class_Initialize()
    Set mcolRemarks = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolRemarks = Nothing
End Sub

